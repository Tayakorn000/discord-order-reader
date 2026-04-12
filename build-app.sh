#!/bin/bash
# สร้าง "Discord Bot.app" ที่กดดับเบิ้ลคลิกได้ใน macOS

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
APP_NAME="Discord Bot"
APP_PATH="$SCRIPT_DIR/$APP_NAME.app"

echo "กำลังสร้าง $APP_NAME.app ..."

# สร้างโครงสร้าง .app bundle
mkdir -p "$APP_PATH/Contents/MacOS"
mkdir -p "$APP_PATH/Contents/Resources"

# สร้าง launcher script ที่อยู่ภายใน .app
cat > "$APP_PATH/Contents/MacOS/launcher.sh" << 'LAUNCHER'
#!/bin/bash
BOT_DIR="$(cd "$(dirname "$0")/../../.." && pwd)"

# เปิด Terminal แล้วรัน bot
osascript << APPLESCRIPT
tell application "Terminal"
    activate
    do script "cd '$BOT_DIR' && echo '=== Discord Order Bot ===' && node index.js"
end tell
APPLESCRIPT
LAUNCHER

chmod +x "$APP_PATH/Contents/MacOS/launcher.sh"

# Info.plist
cat > "$APP_PATH/Contents/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleName</key>
    <string>Discord Bot</string>
    <key>CFBundleDisplayName</key>
    <string>Discord Bot</string>
    <key>CFBundleIdentifier</key>
    <string>com.discordbot.orderreader</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundleExecutable</key>
    <string>launcher.sh</string>
    <key>CFBundleIconFile</key>
    <string>AppIcon</string>
</dict>
</plist>
PLIST

# ดาวน์โหลด icon Discord (ถ้ามี curl)
if command -v curl &>/dev/null; then
    curl -s "https://discord.com/assets/f9bb9c4af2b9c32a2c5ee0014661546d.png" \
        -o "$APP_PATH/Contents/Resources/AppIcon.png" 2>/dev/null || true
fi

echo ""
echo "✅ สร้างเสร็จแล้ว: $APP_PATH"
echo ""
echo "วิธีใช้:"
echo "  ดับเบิ้ลคลิกที่ 'Discord Bot.app' เพื่อเปิดบอท"
echo ""
echo "หมายเหตุ: ต้องมี Node.js ติดตั้งอยู่ในเครื่อง"
echo "ดาวน์โหลด Node.js ได้ที่: https://nodejs.org"
