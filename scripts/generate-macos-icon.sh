#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
ICON_DIR="$ROOT/src-tauri/icons"
mkdir -p "$ICON_DIR"

swift - "$ICON_DIR/icon.png" <<'SWIFT'
import AppKit
import Foundation

let output = CommandLine.arguments[1]
let size = 1024
let image = NSImage(size: NSSize(width: size, height: size))
image.lockFocus()
let bounds = NSRect(x: 28, y: 28, width: 968, height: 968)
let background = NSBezierPath(roundedRect: bounds, xRadius: 218, yRadius: 218)
NSColor(calibratedRed: 0.105, green: 0.105, blue: 0.19, alpha: 1).setFill()
background.fill()
let glow = NSBezierPath(ovalIn: NSRect(x: 120, y: 120, width: 784, height: 784))
NSColor(calibratedRed: 0.39, green: 0.32, blue: 0.92, alpha: 0.24).setFill()
glow.fill()
let attrs: [NSAttributedString.Key: Any] = [
    .font: NSFont.systemFont(ofSize: 560, weight: .heavy),
    .foregroundColor: NSColor(calibratedRed: 0.91, green: 0.89, blue: 1.0, alpha: 1)
]
NSString(string: "T").draw(in: NSRect(x: 290, y: 190, width: 444, height: 640), withAttributes: attrs)
image.unlockFocus()

guard let cgImage = image.cgImage(forProposedRect: nil, context: nil, hints: nil),
      let bitmap = NSBitmapImageRep(
        bitmapDataPlanes: nil,
        pixelsWide: cgImage.width,
        pixelsHigh: cgImage.height,
        bitsPerSample: 8,
        samplesPerPixel: 4,
        hasAlpha: true,
        isPlanar: false,
        colorSpaceName: .deviceRGB,
        bitmapFormat: .alphaNonpremultiplied,
        bytesPerRow: 0,
        bitsPerPixel: 0
      ) else {
    fatalError("Could not create an RGBA bitmap for the TeamImp icon")
}
NSGraphicsContext.saveGraphicsState()
if let context = NSGraphicsContext(bitmapImageRep: bitmap) {
    NSGraphicsContext.current = context
    context.imageInterpolation = .high
    image.draw(in: NSRect(x: 0, y: 0, width: size, height: size))
    context.flushGraphics()
}
NSGraphicsContext.restoreGraphicsState()
guard let png = bitmap.representation(using: .png, properties: [:]) else {
    fatalError("Could not encode the TeamImp RGBA PNG icon")
}
try png.write(to: URL(fileURLWithPath: output))
SWIFT

ICONSET="$ICON_DIR/teamimp.iconset"
rm -rf "$ICONSET"
mkdir -p "$ICONSET"
make_icon() {
  sips -z "$1" "$1" "$ICON_DIR/icon.png" --out "$ICONSET/icon_$2.png" >/dev/null
}
make_icon 16 "16x16"
make_icon 32 "16x16@2x"
make_icon 32 "32x32"
make_icon 64 "32x32@2x"
make_icon 128 "128x128"
make_icon 256 "128x128@2x"
make_icon 256 "256x256"
make_icon 512 "256x256@2x"
make_icon 512 "512x512"
make_icon 1024 "512x512@2x"
iconutil -c icns "$ICONSET" -o "$ICON_DIR/icon.icns"
echo "Generated TeamImp RGBA PNG and ICNS icons."
