# TemplateDependenceChecker
_Version 2.0_

"TemplateDependenceChecker" is a SmartTree enhancement for the Opentext Website Management Server allowing to check the assignment of every template of a content class within all connected projects.

© Stefan Buchali, UDG United Digital Group, www.udg.de

This is third party software. The author is not affiliated in any manner with Open Text Corporation.

## Installation

### Installing the plugin on the server

Copy the folder "TemplateDependenceChecker" into the folder "plugins" of your CMS installation.  
Switch to Server Manager and import the plugin via "Administer Plug-Ins" - "Plug-Ins", action menu "Import Plug-In".  
Assign it to all your desired projects and set the appropriate user rights there if necessary.

The plugin then appears in the action menu of every content class node within these projects.

###  Code settings

#### TemplateDependenceChecker_dlg.asp

This file contains the dialog box the plugin will start with. Translate the german variables (lines 18-28) if necessary.

####  TemplateDependenceChecker_do.asp

This file contains the plugin script itself. Translate the german variables (lines 21-32) if necessary.

The plugin must be run under an administrator account. Please enter the corresponding login data. This user must have the role "Administrator" and "Template 
Editor" in all projects, and the number of allowed sessions should be set to at least 2.

## How to use

In every project, the plugin will appear in the action menu of every content class. After clicking OK, the plugin reads the project variant assignment of the selected content 
class in every project that is connected to the corresponding content class folder.

This process can take some minutes.

Finally it will show a report that lists:

- Every connected project
- And for each one of these projects:
  - Which template of the content class is assigned to a project variant
  - Which one is not assigned at all.
  - The number of active instances of the content class

The report looks like the following:

```
Project: <Name of the connected project>

<Name of the 1st template>: -  
<Name of the 2nd template >: Assigned to <Number> project variants  
<Number> Instances
```

(The "-" after the first template means that it is not assigned).

The following error messages can occur:

- "Insufficient user rights to perform the task" (Variable `dlgInsufficientRights`):  
  The user entered in section [TemplateDependenceChecker_do.asp](#templatedependencechecker_doasp) cannot log in.
- "No rights" (Variable `dlgNoRights`):  
  The user entered in section [TemplateDependenceChecker_do.asp](#templatedependencechecker_doasp) is not assigned to the project.
- "Folder not found" (Variable `dlgFolderNotFound`):  
  The project is connected to the content class folder, but it does not use it.
- "Content class not found" (Variable `dlgContentClassNotFound`):  
  The content class cannot be accessed (e.g. due to authorization packets).

## License and exclusion of liability

This software is licensed under a [Creative Commons GNU General Public License](http://creativecommons.org/licenses/GPL/2.0/). Some rights reserved.

This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 3 of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful, but **without any warranty**; without even the implied warranty of **merchantability** or **fitness for a particular purpose**. See the GNU General Public License for more details.

You should have received a copy of the GNU General Public License along with TemplateDependenceChecker.  If not, see http://www.gnu.org/licenses.

The GNU General Public License is a Free Software license. Like any Free Software license, it grants to you the four following freedoms:

0. The freedom to run the program for any purpose.
1. The freedom to study how the program works and adapt it to your needs.
2. The freedom to redistribute copies so you can help your neighbor.
3. The freedom to improve the program and release your improvements to the public, so that the whole community benefits.

You may exercise the freedoms specified here provided that you comply with the express conditions of this license. The principal conditions are:

- You must conspicuously and appropriately publish on each copy distributed an appropriate copyright notice and disclaimer of warranty and keep intact all the notices that refer to this License and to the absence of any warranty; and give any other recipients of the Program a copy of the GNU General Public License along with the Program. Any translation of the GNU General Public License must be accompanied by the GNU General Public License.
- If you modify your copy or copies of the program or any portion of it, or develop a program based upon it, you may distribute the resulting work provided you do so under the GNU General Public License. Any translation of the GNU General Public License must be accompanied by the GNU General Public License.
- If you copy or distribute the program, you must accompany it with the complete corresponding machine-readable source code or with a written offer, valid for at least three years, to furnish the complete corresponding machine-readable source code.

Any of the above conditions can be waived if you get permission from the copyright holder.

## Changelog

**Version 2.0**  
September 10, 2014  
Add counting instances

**Version 1.2**  
October 25, 2013  
Adaption for version 11 (ENU instead of ENG in plugin XML file)

**Version 1.1**  
April 28, 2011  
Adaption for version 10 (ASP object changed)

**Version 1.0**  
May 6, 2009  
Plugin created