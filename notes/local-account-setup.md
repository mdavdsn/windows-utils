# Set Up Windows With a Local Account

## Current Version

The following instructions are for Windows 11 24H2.

1. Follow the installation until the "Choose a Keyboard Layout" screen.
2. Press <kbd>Shift</kbd> + <kbd>F10</kbd>. A command prompt will appear.
3. Type `OOBE\BYPASSNRO` to disable the internet connection requirement.
4. (Note: only do this step if you cannot or don't want to physically disconnect the somputer from the internet) After the computer reboots, type <kbd>Shift</kbd> + <kbd>F10</kbd> again to open the command prompt. Type`ipconfig /release` to disconnect the internet.
5. Continue with the installation. At the network connection screen, click "I don't have internet" to set up a local username and password.

## Future Versions

Build 2610 and higher of Windows 11 has `OOBE\BYPASSNRO` disabled. Follow the instructions to set up a local account for these versions.

1. Follow the installation until the "Microsoft Experience" sign in screen.
2. Press <kbd>Shift</kbd> + <kbd>F10</kbd>. A command prompt will appear.
3. Type `start ms-cxh:localonly` and a "Create a user for the PC" window will appear.
4. Enter the username and password and complete the installation.

## Windows 11 Install Disk that Bypasses Microsoft Account

The following instructions are to create an installation disk that has the local account setup already activated.

1. Download the Windwos 11 ISO file.
2. Insert a USB flash drive to use as the installation media.
3. Use Rufus and elect the USB drive to write and the ISO file that was downloaded.
4. Click start. A dialog box appears with a few options.
5. Toggle "Remove requirement for an online Microsoft account" to `On` and click OK. There are other options that can be toggled to tweak the installation.
6. Click OK to destroy data on the USB drive and the drive will be written.
7. Follow the installation until the "Choose a Keyboard Layout" screen.
8. Press <kbd>Shift</kbd> + <kbd>F10</kbd>. A command prompt will appear.
9. Type`ipconfig /release` to disconnect the internet.
10. Continue with the installation. At the network connection screen, click "I don't have internet" to set up a local username and password.
