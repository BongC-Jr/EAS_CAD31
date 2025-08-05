## KiWi988 AutoCAD [Productivity] Commands
Developer: **Engr. Bernardo Cabebe Jr.**

## NeuroNet AI Innovations

### Exclusive License and Copyright Notice

***Copyright (c) NeuroNet AI Innovations Inc. All rights reserved.***

This software, including but not limited to all code, patterns, and flow programs, is the exclusive and proprietary property of NeuroNet AI Innovations Inc. All rights, title, and interest in and to this software are owned solely by ***NeuroNet AI Innovations Inc.***

**Retroactive Application**

This exclusive license and copyright notice supersedes any previous licenses, including but not limited to the MIT license, under which this software or any part thereof may have been distributed. All previous releases of this software are now covered by this exclusive license.

**License Terms**

1. Grant of License: NeuroNet AI Innovations Inc. grants you a non-transferable, non-sublicensable license to use this software.
2. Purpose:
    - If NeuroNet AI Innovations Inc. issues an expressed written purpose for the use of this software, you may use it solely for that specified purpose.
    - In the absence of an expressed written purpose, the software may be used for development only and for non-commercial purposes.
3. Restrictions: You may not:
    - Reproduce, modify, or distribute this software or any part thereof without explicit written permission from NeuroNet AI Innovations Inc.
    - Use this software for commercial purposes without a written agreement with NeuroNet AI Innovations Inc.
    - Disclose or reveal any confidential information or trade secrets contained in this software.
4. Ownership: NeuroNet AI Innovations Inc. retains all ownership rights, including but not limited to copyrights, patents, trademarks, and trade secrets, in and to this software.
5. Term and Termination: This license is effective until terminated by NeuroNet AI Innovations Inc. Upon termination, you must cease all use and distribution of this software.

**Disclaimer and Limitation of Liability**

This software is provided "as is" without warranty of any kind. NeuroNet AI Innovations Inc. disclaims all liability for any damages or losses arising from the use of this software.

**Governing Law**

This license and copyright notice shall be governed by and construed in accordance with the laws of Republic of the Philippines, without regard to its conflict of laws principles.

By using this software, you acknowledge that you have read, understood, and agree to be bound by these terms.

NeuroNet AI Innovations Inc. reserves the right to modify or terminate this license at any time.

**Notice**

All copies of this software must include this copyright notice and license terms.

**Action Required**

If you have previously obtained a copy of this software under the MIT license or any other license, you are hereby notified that those licenses are revoked and replaced by this exclusive license. You must comply with the terms of this license for all future use of the software.

----
### Implementation of Sheets V4 in C# .NET [^1]
Steps to perform CRUD Operations on a Google Spreadsheet: 
+ Create a Spreadsheet on Google Sheets 
+ Create Project on Google Developer Console 
+ Enable Google Sheets API & Create Credentials File 
+ Create a .Net Core Project 
+ Install Google Sheets Library using NuGet
  + Install the Google.Apis.Sheets.v4 NuGet package: Install-Package Google.Apis.Sheets.v4 F. Perform Read Write & Update Operations


Setting Up Google Sheet

1. Create Google Spreadsheet and sharing as public
2. Go to Google APIs Console:  https://console.cloud.google.com/apis/dashboard
3. Create a Project. Indicate Project name* Dot Tutorials and Location* No organization
4. After Creating Project you need to Search & Enable Google Sheets Library in your Google API
5. After Enabling Sheets API click on the Create Credentials Button 

[^1]: https://code-maze.com/google-sheets-api-with-net-core/

----
Git Commit
```sh
$ git add file or git add --all
$ git branch -M main
$ git commit -m "Commit 21 Apr 2025 10:46pm BAC"
$ git push -u origin main
```
Git Branch
```sh
$ Git branch nameofbranch
$ Git checkout nameofbranch
$ git add filename "or" git add --all
$ git commit -m "Commit 21 Apr 2025 10:46pm BAC"
$ git push origin nameofbranch
```

Pull from Github
```sh
git branch -a
git checkout branchName
git pull
```

Merge to a main
1. check if you are in main
   ```sh
   $ git branch -a
   ```
2. Go to main
   ```sh
   $ git checkout main
   ```
3. Merge the branch with the main
   ```sh
   $ git merge feature-branch-name
   ```
4. Optionally delete the branch that has been merged

   4.1 Check if the merged branch
      ```sh
      $ git branch --merged
      ```
   4.2 Delete the local branch
      ```sh
      $ git branch -d <branch_name>
      ```
   4.3 Delete the remote branch
      ```sh
      git push origin --delete <branch_name>
      ```
5. Implement **Git commit**
