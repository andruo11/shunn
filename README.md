# shunn
A Microsoft Word add-in to reformat short story text into a publisher-ready manuscript, using the format at https://www.shunn.net/format/story.html. Includes Windows installer files for easy set-up.

This [VSTO add-in](https://en.wikipedia.org/wiki/Visual_Studio_Tools_for_Office) project uses Visual Basic.NET to add a new set of buttons, named "Shunn format," to the **Review** pane in the Office ribbon control at the top of the Word environment. The left-hand button opens a panel where the writer of a short story can update their user information. When the user clicks on the **Create document** button after highlighting the text of their story in the Word editor, a new document will be generated according to the commonly-used editor's style found at the above hyperlink. This style uses different headers on the first page and on the ones after that which are filled in with the user information. 

This is important because editors at busy publishing companies will ignore the short story submissions which don't follow their publishing guidelines. The one used here double-spaces your story to make room for written comments when printed, sets the paragraphs' first-line tab indent, underlines all italicized text, and makes the font 12 pt. Courier. It also adds headers to every page in case your printout is knocked off the editor's desk, and gives the front page of the story a layout they will find familiar. Finally, it replaces all em dashes with two hyphens. Any paragraphs with a style in Word whose name starts with "Head" will be treated as a section header (not in the Shunn format description at the above link) and made bold on a non-indented paragraph.

## Software requirements
Windows & Word 2013 or newer

## Installation files
Download the right file for your version of Windows (103 mb):

[installer_32bit.zip](../../releases/download/1.0/installer_32bit.zip)

[installer_64bit.zip](../../releases/download/1.0/installer_64bit.zip)

Unzip the file and double-click on setup.exe to install. 

Includes .NET and VSTO runtime prerequisites for users who don't already have them installed from another program, which explains the large file size for just an add-in.

### Project build environment:
Visual Studio Community 2017 + Installer extension at https://marketplace.visualstudio.com/items?itemName=VisualStudioClient.MicrosoftVisualStudio2017InstallerProjects
