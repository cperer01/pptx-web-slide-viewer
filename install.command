#!/bin/bash
# Web Slide installer for PowerPoint on Mac.
# Double-click this file to install. It copies the add-in manifest into the
# location PowerPoint reads add-ins from. No Terminal knowledge needed.

DIR="$(cd "$(dirname "$0")" && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef"

if [ ! -f "$DIR/manifest.xml" ]; then
  osascript -e 'display dialog "manifest.xml was not found next to this installer. Keep both files in the same folder and try again." buttons {"OK"} default button "OK" with icon stop' >/dev/null 2>&1
  exit 1
fi

mkdir -p "$WEF"
cp "$DIR/manifest.xml" "$WEF/"

osascript -e 'display dialog "Web Slide is installed.\n\nQuit PowerPoint if it is open, reopen it, then go to Insert > Add-ins > My Add-ins and choose Web Slide." buttons {"OK"} default button "OK"' >/dev/null 2>&1
echo "Installed manifest to: $WEF"
