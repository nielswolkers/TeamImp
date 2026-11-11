use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use keyring::Entry;
use rand::{rngs::OsRng, RngCore};
use serde_json::Value;
use sha2::{Digest, Sha256};
use std::{
    collections::HashMap,
    io::{Read, Write},
    net::TcpListener,
    process::Command,
    thread,
    time::{Duration, Instant},
};
use tauri::Emitter;
use url::Url;

const SERVICE: &str = "com.nielswolkers.teamimp";
const ACCOUNT: &str = "microsoft-refresh-token";
const BASE_SCOPES: &[&str] = &[
    "openid",
    "profile",
    "offline_access",
    "User.Read",
    "EduAssignments.ReadBasic",
    "Chat.ReadBasic",
    "Chat.Read",
];

#[tauri::command]
fn start_microsoft_login(
    app: tauri::AppHandle,
    client_id: String,
    extra_scopes: Vec<String>,
) -> Result<(), String> {
    let client_id = client_id.trim().to_string();
    if client_id.is_empty() {
        return Err("Microsoft-login is nog niet geconfigureerd voor deze TeamImp-build. De publieke Microsoft Application ID ontbreekt.".into());
    }

    let listener = TcpListener::bind("localhost:0").map_err(|e| format!("Lokale aanmeldpoort openen mislukt: {e}"))?;
    listener.set_nonblocking(true).map_err(|e| e.to_string())?;
    let port = listener.local_addr().map_err(|e| e.to_string())?.port();
    let redirect_uri = format!("http://localhost:{port}");

    let mut verifier_bytes = [0_u8; 32];
    let mut state_bytes = [0_u8; 32];
    OsRng.fill_bytes(&mut verifier_bytes);
    OsRng.fill_bytes(&mut state_bytes);
    let verifier = URL_SAFE_NO_PAD.encode(verifier_bytes);
    let state = URL_SAFE_NO_PAD.encode(state_bytes);
    let challenge = URL_SAFE_NO_PAD.encode(Sha256::digest(verifier.as_bytes()));

    let mut scopes: Vec<String> = BASE_SCOPES.iter().map(|s| (*s).to_string()).collect();
    for scope in extra_scopes {
        if !scopes.contains(&scope) {
            scopes.push(scope);
        }
    }
    let scope_string = scopes.join(" ");

    let mut query = url::form_urlencoded::Serializer::new(String::new());
    query
        .append_pair("client_id", &client_id)
        .append_pair("response_type", "code")
        .append_pair("redirect_uri", &redirect_uri)
        .append_pair("response_mode", "query")
        .append_pair("scope", &scope_string)
        .append_pair("state", &state)
        .append_pair("code_challenge", &challenge)
        .append_pair("code_challenge_method", "S256")
        .append_pair("prompt", "select_account");
    let authorize_url = format!(
        "https://login.microsoftonline.com/common/oauth2/v2.0/authorize?{}",
        query.finish()
    );

    thread::spawn(move || {
        let result = complete_browser_login(
            listener,
            &client_id,
            &redirect_uri,
            &verifier,
            &state,
            &scope_string,
            &authorize_url,
        );
        match result {
            Ok((access_token, expires_in)) => {
                let _ = app.emit(
                    "microsoft-auth-success",
                    serde_json::json!({
                        "access_token": access_token,
                        "expires_in": expires_in
                    }),
                );
            }
            Err(error) => {
                let _ = app.emit("microsoft-auth-error", error);
            }
        }
    });

    Ok(())
}

fn complete_browser_login(
    listener: TcpListener,
    client_id: &str,
    redirect_uri: &str,
    verifier: &str,
    expected_state: &str,
    scope: &str,
    authorize_url: &str,
) -> Result<(String, u64), String> {
    open_system_browser(authorize_url)?;

    let deadline = Instant::now() + Duration::from_secs(300);
    let (mut stream, _) = loop {
        match listener.accept() {
            Ok(connection) => break connection,
            Err(error) if error.kind() == std::io::ErrorKind::WouldBlock => {
                if Instant::now() >= deadline {
                    return Err("Microsoft-aanmelding is verlopen. Probeer opnieuw.".into());
                }
                thread::sleep(Duration::from_millis(150));
            }
            Err(error) => return Err(format!("Microsoft-redirect ontvangen mislukt: {error}")),
        }
    };

    stream.set_read_timeout(Some(Duration::from_secs(10))).map_err(|e| e.to_string())?;
    let mut buffer = [0_u8; 16_384];
    let length = stream.read(&mut buffer).map_err(|e| format!("Microsoft-antwoord lezen mislukt: {e}"))?;
    let request = String::from_utf8_lossy(&buffer[..length]);
    let request_target = request
        .lines()
        .next()
        .and_then(|line| line.split_whitespace().nth(1))
        .ok_or_else(|| "Ongeldig antwoord van Microsoft ontvangen.".to_string())?;
    let callback = Url::parse(&format!("http://localhost{request_target}"))
        .map_err(|e| format!("Microsoft-redirect verwerken mislukt: {e}"))?;
    let parameters: HashMap<String, String> = callback.query_pairs()
        .map(|(key, value)| (key.into_owned(), value.into_owned()))
        .collect();

    let response_body = "<!doctype html><html><head><meta charset=\"utf-8\"><title>TeamImp</title></head><body style=\"font-family:system-ui;text-align:center;padding:64px\"><h2>Microsoft-aanmelding verwerkt</h2><p>Je kunt dit tabblad sluiten en teruggaan naar TeamImp.</p></body></html>";
    let response = format!(
        "HTTP/1.1 200 OK\r\nContent-Type: text/html; charset=utf-8\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{}",
        response_body.len(),
        response_body
    );
    let _ = stream.write_all(response.as_bytes());

    if let Some(error) = parameters.get("error") {
        let description = parameters.get("error_description").cloned().unwrap_or_else(|| error.clone());
        return Err(format!("Microsoft-aanmelding geweigerd: {description}"));
    }
    if parameters.get("state").map(String::as_str) != Some(expected_state) {
        return Err("Microsoft-aanmelding ongeldig: de beveiligingscontrole (state) komt niet overeen.".into());
    }
    let code = parameters.get("code").ok_or_else(|| "Microsoft heeft geen autorisatiecode teruggestuurd.".to_string())?;

    let response = reqwest::blocking::Client::new()
        .post("https://login.microsoftonline.com/common/oauth2/v2.0/token")
        .form(&[
            ("client_id", client_id),
            ("grant_type", "authorization_code"),
            ("code", code.as_str()),
            ("redirect_uri", redirect_uri),
            ("code_verifier", verifier),
            ("scope", scope),
        ])
        .send()
        .map_err(|e| format!("Verbinding met Microsoft-tokenservice mislukt: {e}"))?;

    let status = response.status();
    let payload: Value = response.json().map_err(|e| format!("Microsoft-tokenantwoord kon niet worden gelezen: {e}"))?;
    if !status.is_success() {
        let description = payload["error_description"].as_str()
            .or_else(|| payload["error"].as_str())
            .unwrap_or("Onbekende Microsoft OAuth-fout");
        return Err(format!("Microsoft-login ({status}): {description}"));
    }

    let access_token = payload["access_token"].as_str()
        .ok_or_else(|| "Microsoft heeft geen access token teruggestuurd.".to_string())?
        .to_string();
    if let Some(refresh_token) = payload["refresh_token"].as_str() {
        let entry = Entry::new(SERVICE, ACCOUNT).map_err(|e| format!("macOS Keychain openen mislukt: {e}"))?;
        entry.set_password(refresh_token).map_err(|e| format!("Microsoft-token veilig opslaan mislukt: {e}"))?;
    }
    let expires_in = payload["expires_in"].as_u64().unwrap_or(3600);
    Ok((access_token, expires_in))
}

fn open_system_browser(url: &str) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    let result = Command::new("open").arg(url).spawn();

    #[cfg(target_os = "windows")]
    let result = Command::new("cmd").args(["/C", "start", "", url]).spawn();

    #[cfg(target_os = "linux")]
    let result = Command::new("xdg-open").arg(url).spawn();

    result.map(|_| ()).map_err(|e| format!("Microsoft-aanmeldpagina openen mislukt: {e}"))
}

#[tauri::command]
async fn refresh_microsoft_access_token(
    client_id: String,
    extra_scopes: Vec<String>,
) -> Result<Value, String> {
    let client_id = client_id.trim().to_string();
    if client_id.is_empty() {
        return Err("Microsoft Application ID ontbreekt in deze TeamImp-build.".into());
    }
    let refresh_token = {
        let entry = Entry::new(SERVICE, ACCOUNT).map_err(|e| e.to_string())?;
        entry.get_password()
            .map_err(|e| format!("Microsoft-sessie ophalen uit macOS Keychain mislukt: {e}"))?
    };
    let mut scopes: Vec<String> = BASE_SCOPES.iter().map(|s| (*s).to_string()).collect();
    for scope in extra_scopes {
        if !scopes.contains(&scope) {
            scopes.push(scope);
        }
    }
    let scope = scopes.join(" ");
    let response = reqwest::Client::new()
        .post("https://login.microsoftonline.com/common/oauth2/v2.0/token")
        .form(&[
            ("client_id", client_id.as_str()),
            ("grant_type", "refresh_token"),
            ("refresh_token", refresh_token.as_str()),
            ("scope", scope.as_str()),
        ])
        .send()
        .await
        .map_err(|e| format!("Microsoft-sessie vernieuwen mislukt: {e}"))?;
    let status = response.status();
    let payload: Value = response.json().await
        .map_err(|e| format!("Microsoft-tokenantwoord kon niet worden gelezen: {e}"))?;
    if !status.is_success() {
        let description = payload["error_description"].as_str()
            .or_else(|| payload["error"].as_str())
            .unwrap_or("Onbekende Microsoft OAuth-fout");
        return Err(format!("Microsoft-token vernieuwen ({status}): {description}"));
    }
    if let Some(new_refresh_token) = payload["refresh_token"].as_str() {
        let entry = Entry::new(SERVICE, ACCOUNT).map_err(|e| e.to_string())?;
        entry.set_password(new_refresh_token)
            .map_err(|e| format!("Microsoft-sessie veilig opslaan mislukt: {e}"))?;
    }
    Ok(serde_json::json!({
        "access_token": payload["access_token"].as_str()
            .ok_or_else(|| "Microsoft heeft geen access token teruggestuurd.".to_string())?,
        "expires_in": payload["expires_in"].as_u64().unwrap_or(3600)
    }))
}

#[tauri::command]
fn get_microsoft_refresh_token() -> Result<Option<String>, String> {
    let entry = Entry::new(SERVICE, ACCOUNT).map_err(|e| e.to_string())?;
    match entry.get_password() {
        Ok(token) => Ok(Some(token)),
        Err(keyring::Error::NoEntry) => Ok(None),
        Err(error) => Err(error.to_string()),
    }
}

#[tauri::command]
fn set_microsoft_refresh_token(token: String) -> Result<(), String> {
    let entry = Entry::new(SERVICE, ACCOUNT).map_err(|e| e.to_string())?;
    entry.set_password(&token).map_err(|e| e.to_string())
}

#[tauri::command]
fn delete_microsoft_refresh_token() -> Result<(), String> {
    let entry = Entry::new(SERVICE, ACCOUNT).map_err(|e| e.to_string())?;
    match entry.delete_credential() {
        Ok(()) | Err(keyring::Error::NoEntry) => Ok(()),
        Err(error) => Err(error.to_string()),
    }
}

pub fn run() {
    tauri::Builder::default()
        .invoke_handler(tauri::generate_handler![
            start_microsoft_login,
            refresh_microsoft_access_token,
            get_microsoft_refresh_token,
            set_microsoft_refresh_token,
            delete_microsoft_refresh_token
        ])
        .run(tauri::generate_context!())
        .expect("error while running TeamImp");
}
