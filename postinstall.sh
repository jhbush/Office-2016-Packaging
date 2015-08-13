#!/bin/bash

# Script to install Office 2016 and stop the annoying AU Daemon first run message
# Based on work by https://derflounder.wordpress.com/2015/08/05/creating-an-office-2016-15-12-3-installer/
# and http://macops.ca/disabling-first-run-dialogs-in-office-2016-for-mac/
# And on work by http://www.officeformachelp.com/ (the removal script for office 2011)
# Massive amounts of work also done by the MANY admins on Slack MacAdmin #microsoft-office

# Author : contact@richard-purves.com

# Set up log file, folder and function
LOGFOLDER="/private/var/log/organisation name here"
LOG=$LOGFOLDER"/Office-2016-Install.log"
error=0

if [ ! -d "$LOGFOLDER" ];
then
mkdir $LOGFOLDER
fi

logme()
{
# Check to see if function has been called correctly
if [ -z "$1" ]
then
echo $( date )" - logme function call error: no text passed to function! Please recheck code!"
exit 1
fi

# Log the passed details
echo $( date )" - "$1 >> $LOG
echo "" >> $LOG
}

# Office 2011 Install location

office2011="/Applications/Microsoft Office 2011"
logme "Checking if Office 2011 is present"

# If installed, then clean up files

if [ -d "$office2011" ];
then
logme "Office 2011 installation detected. Removing."

# Stop Office 2011 background processes
logme "Stopping Office 2011 background processes"
osascript -e 'tell application "Microsoft Database Daemon" to quit' | tee -a ${LOG}
osascript -e 'tell application "Microsoft AU Daemon" to quit' | tee -a ${LOG}
osascript -e 'tell application "Office365Service" to quit' | tee -a ${LOG}

# Delete external applications apart from Lync
logme "Deleting Office 2011 applications"
rm -R '/Applications/Microsoft Communicator.app/' | tee -a ${LOG}
rm -R '/Applications/Microsoft Messenger.app/' | tee -a ${LOG}
rm -R '/Applications/Microsoft Office 2011/' | tee -a ${LOG}
rm -R '/Applications/Remote Desktop Connection.app/' | tee -a ${LOG}

# Delete MS working folder
logme "Deleting /Library/Application Support/Microsoft"
rm -R '/Library/Application Support/Microsoft/' | tee -a ${LOG}

# Remove all Automator actions
logme "Deleting Automator actions"
rm -R /Library/Automator/*Excel* | tee -a ${LOG}
rm -R /Library/Automator/*Office* | tee -a ${LOG}
rm -R /Library/Automator/*Outlook* | tee -a ${LOG}
rm -R /Library/Automator/*PowerPoint* | tee -a ${LOG}
rm -R /Library/Automator/*Word* | tee -a ${LOG}
rm -R /Library/Automator/*Workbook* | tee -a ${LOG}
rm -R '/Library/Automator/Get Parent Presentations of Slides.action' | tee -a ${LOG}
rm -R '/Library/Automator/Set Document Settings.action' | tee -a ${LOG}

# Remove Office Fonts and copy disabled ones back into place
logme "Deleting Microsoft Fonts folder"
rm -R /Library/Fonts/Microsoft/ | tee -a ${LOG}

logme "Moving previously disabled fonts back to main fonts folder"
mv '/Library/Fonts Disabled/Arial Bold Italic.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Arial Bold.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Arial Italic.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Arial.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Brush Script.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Times New Roman Bold Italic.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Times New Roman Bold.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Times New Roman Italic.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Times New Roman.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Verdana Bold Italic.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Verdana Bold.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Verdana Italic.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Verdana.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Wingdings 2.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Wingdings 3.ttf' /Library/Fonts | tee -a ${LOG}
mv '/Library/Fonts Disabled/Wingdings.ttf' /Library/Fonts | tee -a ${LOG}

# Remove Sharepoint plugin
logme "Deleting Sharepoint folder"
rm -R /Library/Internet\ Plug-Ins/SharePoint* | tee -a ${LOG}

# Finally remove LaunchDaemons, preference files and any helper tools
logme "Deleting LaunchDaemons, Prefs and helper tools"
rm -R /Library/LaunchDaemons/com.microsoft.* | tee -a ${LOG}
rm -R /Library/Preferences/com.microsoft.* | tee -a ${LOG}
rm -R /Library/PrivilegedHelperTools/com.microsoft.* | tee -a ${LOG}
else
logme "Office 2011 not installed. Skipping uninstallation."
fi

logme "Starting Installation of Microsoft Office 2016"

# Determine working directory

install_dir=`dirname $0`
logme "Working Directory: $install_dir"

# Install all the packages

logme "Installing Microsoft Excel 2016"
/usr/sbin/installer -dumplog -verbose -pkg $install_dir/"Microsoft_Excel.pkg" -target "$3" | tee -a ${LOG}

logme "Installing Microsoft OneNote 2016"
/usr/sbin/installer -dumplog -verbose -pkg $install_dir/"Microsoft_OneNote.pkg" -target "$3" | tee -a ${LOG}

logme "Installing Microsoft Outlook 2016"
/usr/sbin/installer -dumplog -verbose -pkg $install_dir/"Microsoft_Outlook.pkg" -target "$3" | tee -a ${LOG}

logme "Installing Microsoft PowerPoint 2016"
/usr/sbin/installer -dumplog -verbose -pkg $install_dir/"Microsoft_PowerPoint.pkg" -target "$3" | tee -a ${LOG}

logme "Installing Microsoft Word 2016"
/usr/sbin/installer -dumplog -verbose -pkg $install_dir/"Microsoft_Word.pkg" -target "$3" | tee -a ${LOG}

# Register Microsoft AU Daemon for all users

# Set up variables

register_trusted_cmd="/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -R -f -trusted"
application="/Library/Application Support/Microsoft/MAU2.0/Microsoft AutoUpdate.app/Contents/MacOS/Microsoft AU Daemon.app"

logme "Registering the Microsoft Auto Update Daemon"
$register_trusted_cmd "$application" | tee -a ${LOG}

# Let's make sure the first run prompting is disabled where possible.

logme "Disabling the application first run prompting"
defaults write /Library/Preferences/com.microsoft.Excel kSubUIAppCompletedFirstRunSetup1507 -bool true | tee -a ${LOG}
defaults write /Library/Preferences/com.microsoft.onenote.mac kSubUIAppCompletedFirstRunSetup1507 -bool true | tee -a ${LOG}
defaults write /Library/Preferences/com.microsoft.Outlook kSubUIAppCompletedFirstRunSetup1507 -bool true | tee -a ${LOG}
defaults write /Library/Preferences/com.microsoft.Outlook FirstRunExperienceCompletedO15 -bool true | tee -a ${LOG}
defaults write /Library/Preferences/com.microsoft.PowerPoint kSubUIAppCompletedFirstRunSetup1507 -bool true | tee -a ${LOG}
defaults write /Library/Preferences/com.microsoft.Word kSubUIAppCompletedFirstRunSetup1507 -bool true | tee -a ${LOG}

logme "Package installation completed."
exit 0
