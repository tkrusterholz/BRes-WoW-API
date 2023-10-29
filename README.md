# BRes-WoW-API
GAS script for polling the WoW Character API, retrieving character information, and pasting into a character spreadsheet.

## Important Notes
- Script only works for one realm, and defaults to Shadowsong -- see "Personalize The Script" below to change this
- The entire script will fail if API returns a 404 ('character not found') for any call.
	- Most common causes of a 404:
		- Your character name is misspelled in your spreadsheet - double check any fancy characters!
		- Your character is not in the API database (i.e. under level 10)
			- I don't know how long it takes for this database to update after you hit level 10, but in testing it seems to take over 24 hours
			- You can check by going to worldofwarcraft.com and searching for your character, if they appear in search results then they should appear in the API
	- Best way to mitigate this (for now) is to just delete characters under level 10 from the spreadsheet, or move them out of the character name cell range
	- In future I will make the script capable of ignoring 404s instead of halting

## How To Make The Script Work For You
Follow ALL of these steps.  If you do not do this correctly, I am not responsible for any loss of very very critical character spreadsheet data!
### Create The Script File
1. Open your spreadsheet in a web browser.  It MUST be hosted in Google Sheets for any of this to work!
2. In the menu bar, choose Extensions > Apps Script, your browser will open a script editor in a new tab
3. Delete all the text in the editor so the file is blank
4. Copy the contents of the Update Script (found [here](https://github.com/tkrusterholz/BRes-WoW-API/blob/main/UpdateScript.gs)) and paste into the script editor
5. In the script editor menu bar, click the Save icon and verify there are no errors
### Add The Menu Bar Item
1. Make sure you have your spreadsheet open to the correct sheet in another tab (if you just did the above section then you are set)
2. In the menu bar of the script editor, find the drop down where you can choose a function and choose "addMenuItem," then click Run
3. A pop-up will appear asking for authorization, this is granting the script read/write access to your spreadsheet
4. Click Review Permissions, then choose your correct Google account, then you will see a "Google hasn't verified this app" warning, click "Advanced," then "Go to Untitled project," then "Allow"
5. You should see the Execution log appear and post "Execution completed," go back to your spreadsheet and verify you now see a menu bar item "WoW API" with option "Update Character Info" - do not actually run it yet!!
### Add The Library
1. You will need access to a script ("library") that I have created, send me the email address of your Google account that owns your spreadsheet
2. I will add your account to my whitelist, and then respond with a "Script ID"
3. In the script editor window, on the left side find "Libraries" and then click the + to add a new library
4. Paste in the script ID, click "Look up", you should see BResWoWAPIToken
5. Make sure you choose the latest version (probably version 2) and then click 'Add'
### Personalize The Script
1. Look at the top of the script for the section "function personalize()"
2. Underneath that are four lines with comment labels, you need to update these with your preferred Realm name and the correct cell ranges on your spreadsheet
3. Once the realm and cell range values are correct for your spreadsheet, Save the script again
### Make Magic Happen
1. Go back to your spreadsheet, from the menu bar choose WoW API > Update Character Info
2. You will get another request to authorize spreadsheet access, follow the same steps as before
3. Cross your fingers neither of us messed anything up and it works the first time!
4. If something goes wrong and you overwrite the wrong cells, simply Edit > Undo in the spreadsheet
