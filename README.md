![PowerShell](https://img.shields.io/badge/powershell-5391FE?style=flat&logo=powershell&logoColor=white)&nbsp;&nbsp;&nbsp;[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-green)](https://www.gnu.org/licenses/gpl-3.0)

# UpdateQueryGroup
A declarative way to implement dynamic groups in Active Directory.

## Motivation
Since its introduction in the year 2000, Active Directory has not received a feature for dynamic group memberships comparable to that of Novell's directory service NDS (aka eDirectory). While there are isolated approaches from Microsoft, for example, for Microsoft Exchange with dynamic distribution lists based on PowerShell or in Azure with dynamic members, using a dynamic group in on-premises AD still equires third-party tools. The author of this script had his own ideas about what a dynamic group should be able to do and is well aware that there are probably already numerous other attempts, even better ones, to achieve the same thing. Nevertheless, he deliberately did not inform himself about other solutions before starting this project in order not to be discouraged, and the result of this ignorance is now presented here.


## Installation
The installation only consists of copying the script files into a desired folder and possibly creating one or more scheduled tasks. For each instance the Config.json file must be configured.

## Description
The solution presented here has the following features that the author found useful: 

+   Mapping OUs into groups and vice versa.
+	Mapping groups into groups.
+   Source groups can be linked with the logical operators AND or OR.
+   LDAP Filter.
+   PowerShell Expression Language Filter.
+   Transformation of samAccountNames.
+   Declarative definition directly at the group object.
+   Exclude rules based on regular expressions.
+   Distribute group updates across different instances/servers.
+   WhatIf mode to protect the productive environment.
+   Log of all dynamic group changes.
+   Aliases for LDAP attribute names to keep the declaration text compact.
+   Container paths as short as possible to keep the declaration text compact.


<img width="408" height="482" alt="DynamicGroup_Example" src="https://github.com/user-attachments/assets/75bdef34-27b3-4b79-93eb-9a06d14fcc2c" />


For a detailed description and usage, the interested reader should please study the about_UpdateQueryGroup file, otherwise everything would have to be repeated here.

## Contributing
All PowerShell developers or Active Directory experts are very welcome to help and make the code better, more readable or contribute new ideas. 


## License

This project is licensed under the terms of the GPL V3 license. Please see the included LICENCE file gor more details.

## Release History

### Version 0.1.0 (2025/07/22)
First release, testing has been done but bugs may still exist.

