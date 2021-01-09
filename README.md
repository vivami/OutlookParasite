# OutlookParasite

This is a method that misuses Outlook Add-in functionality to obtain (unprivileged) persistence using Outlook (or other Office programs). This method also bypasses the "ClickOnce" install pop-up that you'd normally get when installing an unsigned Outlook Add-in. This is pretty stealth I guess, since you're living inside an Outlook process and are started once Outlook is started by the user (every morning?). It's also not detected by Sysinternals' Autoruns. More information [here](https://vanmieghem.io/stealth-outlook-persistence/).

## Usage

1. Compile the `.sln` and copy everything in the `Release` directory except for the `.pdo` to the target machine in some directory (i.e. `C:\ProgramData\`).
2. Execute the `Install-OutlookAddin -PayloadPath C:\ProgramData\<OutlookAddinNameYouCanChooseYourself>.vsto`

To clean up, run `Remove-OutlookAddin` and delete the files on disk.
