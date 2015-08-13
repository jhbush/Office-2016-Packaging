#!/bin/bash

# Uninstallation script for Office 2016 installed by my method ;)

# Author : contact@richard-purves.com

# With anything that deletes stuff, TEST BEFORE PUTTING INTO PRODUCTION!

# Delete apps from system

rm -rf /Applications/Microsoft\ Word.app
rm -rf /Applications/Microsoft\ Excel.app
rm -rf /Applications/Microsoft\ PowerPoint.app
rm -rf /Applications/Microsoft\ OneNote.app
rm -rf /Applications/Microsoft\ Outlook.app

# Delete licencing file

rm -rf /Library/Preferences/com.microsoft.office.licensingV2.plist

# Forget the pkg info from the system.

/usr/sbin/pkgutil --forget organisation.pkg.MicrosoftOffice2016

# All done.
