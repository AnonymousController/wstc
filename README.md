# WebScheduleToCalendar
Script to copy a user's work schedule from Web Scheduler to the Calendar application.
<br/>
<br/>
<br/>
<br/>

# How to use
Download the file WebScheduleToCalendar.scpt and open it in Script&nbsp;Editor (a default macOS application found in /Applications/Utilities), or copy the contents of that file and paste into a new Script Editor file.

Save the script anywhere (I usually use ~/Library/Scripts, which may not exist by default) and run it when desired.
You might set it to run on login, or with a `cron` job, or from the macOS menu bar.

Probably the easiest for most people, however, will be to save the script as a standalone application using File&nbsp;>&nbsp;Export… in Script&nbsp;Editor and selecting File&nbsp;Format:&nbsp;Application. Again, the script can be saved anywhere, including the standard /Applications folder.
<br/>

## Requirements
The script assumes you have the applications "Safari" and "Calendar" on your Mac, and have not renamed them.
There must be at least one calendar in Calendar; the script will not create one.
You must allow access to the applications the script requests and must answer the questions it asks, as detailed below.
<br/>

## Script prompts
The script will ask for the following:

#### Finder access
The script needs to read settings files from ~/Library/Application&nbsp;Support (like most other applications), and if the files do not exist it uses Finder to create them.

#### System Events access
The script uses System&nbsp;Events to read information from the settings files and to manipulate the visibility of the Safari and Calendar processes so they do not take over your screen.

#### Calendar access
The script needs to control Calendar in order to open the application.
It also needs access to the calendars on your device in order to create and edit work events.

#### Safari access
The script needs to control Safari in order to scrape information about your shifts from WebScheduler.
It will keep Safari in the same configuration as when the script launched (if it was closed, it will quit; if not, it will remain open).

#### JavaScript access
The script uses JavaScript within Safari to do the bulk of the work.
More information on enabling JavaScript access can be found at the [dedicated instruction page](AllowJS.md).

### User settings
The script will ask several questions regarding settings for the work events.
These should all be self-explanatory.
The facility address is only used to populate the location field for the calendar events, and may be safely left blank.

### Login settings
The script signs in to MyAccess automatically, and therefore needs to store your MyAccess credentials.
The credentials are stored in an encrypted format using a password that is randomly generated when you save the script file.
However, see the "Security" section later for more information and caveats.

When the script reaches a secret question for which it does not have a stored answer, it will ask for your answer.
The next time it encounters the question it will use the answer and you will not have to enter it again.

If MyAccess rejects either your PIN or secret answer as incorrect, the script will wipe out its stored login information and will ask you to re-enter your credentials.
<br/>

## Troubleshooting
The script has some amount of error handling and checking built in, to the point that most of the bugs have been either fixed or will be caught and corrected.
The main problem that I see consistently is: The script moves ahead too fast while a web page has not finished loading.
This can cause several errors, most commonly something along the lines of "*variable* is not defined."
If this occurs, try simply running the script again once or twice; the page in question may load faster and the error will be avoided.

If you have a slow Internet connection, edit the property `globalPreDelay`, which is the very first property defined in the script.
The script was written to compromise between waiting long enough to see if a new page has loaded and not waiting too long if it turns out it hasn't.
Adding a five- or ten-second wait before attempting this check may be enough to keep from failing.

Any errors that are not resolved by increasing `globalPreDelay` should be brought up in this project's Issues section.
<br/>
<br/>
<br/>
<br/>

# Security
This script asks for your MyAccess login information.
Obviously this is serious business, and you should be wary of any third-party script or software that asks for such information.

When deciding whether to use this script, consider the following:
1. **What does the script do with the information?**
For example, does it send me an email so I can go into your Employee Express and change your direct deposit?
You can read through the code yourself and see that it does not.
The information is only used to sign in to MyAccess and is only used while the script is running, no other purpose whatsoever.
2. **How does the script store the information?**
The login information is stored to your computer's hard drive for use next time you run the script.
The information is encrypted with the `openssl` command using AES-256 encryption—your PIN and secret answers are *not* stored in plaintext, so someone looking through your files won't know what your credentials are.
3. **How does the script *use* the information?**
Within the script itself, your PIN and answers are not stored in plaintext as long-term or even as intentional variables.
The plaintext values are immediately passed to the encrypting function and only the encrypted results are stored.
When the script places your plaintext answers into the login page, the plaintext answers are not stored—they are decrypted as the script needs them.

### Caveats
Although the plaintext login credentials are not stored on the disk, they can be gleaned at several points:
* The results of the entries you make can be logged by the Script&nbsp;Editor, and can be stored and reviewed depending on your Script&nbsp;Editor settings.
* The script calls a `bash`/`zsh` shell command, `openssl`, which can be used to encrypt or decrypt a file or string.
In this case, the plaintext credential is passed to the shell command, along with the plaintext random-but-persistent password used as the encryption key.
A record is kept of commands that you yourself run in the Terminal application; I don't believe a similar record is kept of commands run by applications and scripts, but I could be mistaken.
* If you use a script debugging application an even deeper log can be made of all variables, even temporary variables, used in the script.
The random password used for encrypting your information can also be recorded, allowing someone to decrypt the information stored on your hard drive.

In addition, decrypting a user's credentials is *not* best security practice; any good login system will, on the back side, simply take the user's entered credentials and encrypt them, then compare the result with the stored encrypted credentials to see if there is a match.
Of course, there are password manager applications which do a similar thing to what this script does (though they are designed by professional teams and are almost certainly more secure).

Personally I find all of these caveats worth the risk.
I subscribe to the theory that if an attacker has physical, or even remote, access to your computer you've already lost.
But you are welcome to modify the code to suit your own needs, for example by configuring delays while your password manager enters in the information so this script doesn't have to.
<br/>
<br/>
<br/>
<br/>

# Acknowledgements
This script was inspired by the iOS app VoidTime, which was written by someone who will be named if he lets me know he wants to be.



