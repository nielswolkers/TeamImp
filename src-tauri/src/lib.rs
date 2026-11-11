use keyring::Entry;

const SERVICE: &str = "com.nielswolkers.teamimp";
const ACCOUNT: &str = "microsoft-refresh-token";

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
            get_microsoft_refresh_token,
            set_microsoft_refresh_token,
            delete_microsoft_refresh_token
        ])
        .run(tauri::generate_context!())
        .expect("error while running TeamImp");
}
