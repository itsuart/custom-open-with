# My Open with shell extension
Ever wanted full control over contents of the 'Open with' submenu in the File Explorer? This project is the answer to that need!

For example following structure will allow to open .txt files in Emacs and Notepad:
![My Open with txt handlers folder](https://github.com/itsuart/custom-open-with/raw/master/img/txt-folder.png)

As you migh have guessed, contents of (.exe) folder contains handlers for .exe files, (.zip) for .zip files and so one, so forth.
Unlike File Exprorer's though, you are not limited to 3 items per extension, and can have as many as you want. Additionally, our handlers form hierarchy.
For example, a context menu for a .txt file will look like this:
![My Open with txt menu](https://github.com/itsuart/custom-open-with/raw/master/img/txt-menu.png)
Emacs and Notepad came from 'Documents\Open With Handlers for\Files by Extension\(.txt)' subfolder, HxD (a hex editor) from 'All files', and 'Copy path' -- from 'Everything' ('Open handlers folder' is built-in menu item and will open the handlers folder in the explorer). As you can see, most specific handlers are at the top, and less specific ones -- at the bottom.

# Handler folders
`All files` -- these handlers could be used for any file.
`Everything` -- these handlers colud be used for anything, be it a file or a folder.
`Extensionless files` -- these handlers could be used to open files that don't have extension (ex: LICENSE, COPYING).
`Files by extension\(.extension)` -- these handlers could be used to open files, that have `.extension` extension.
`Folders` -- these handlers could be used with folders.

# How to use
* Navigate to Release tab to get prebuilt version of the extension and (un)installer.
* Unpack the zip file somewhere (I'm using 'c:\tools\my open with' for example).
* Run installer.exe to create folders and register the extension.
* Populate installer-created folders with links, executables, or .cmd files, and they will be called for selected filesystem objects (folders and files).


# How to uninstall
Run the install.exe again.