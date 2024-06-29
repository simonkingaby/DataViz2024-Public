# GitLab, GitHub, Git Bash
Today, we're going to set up a Git environment for all the code repositories you will use in class. It's complex and confusing, but I hope we can get it sorted all out today.

Here is an overview of the steps:
1) We'll create a parent folder named DataViz2024 to contain everything.
2) We'll create an SSH Key and Clone the GitLab repo.
3) We'll copy the GitLab Class Activities to a Working Folder where we can make changes to the files.
4) We'll log in to GitHub and create a Personal Access Token (PAT).
5) We'll clone Simon's supplemental class repo, which has a few useful things in it.
6) We'll create public repos for the Homework files.
7) We'll clone those Homework repos into our local DataViz2024 folder.
8) We'll take a break to celebrate because that's a job well done.
The following sections will provide step by step instructions for the steps above.

## 1. Create the DataViz2024 folder
Follow these steps:
1) Right-click the Desktop and create a New Folder
2) Name it DataViz2024

## 2. Create an SSH Key and Clone the GitLab Repo
Follow the steps found here:
[UNC Chapel Hill / UNC-VIRT-DATA-PT-06-2024-U-LOLC · GitLab](https://git.bootcampcontent.com/UNC-Chapel-Hill/UNC-VIRT-DATA-PT-06-2024-U-LOLC)

## 3. Create a Working Folder for Class Activities
Follow these steps:
1) In the DataViz2024 folder, create a New Folder called WorkingFolder
2) Now using File Explorer (PC) or Finder (Mac) go to the UNC-VIRT folder you just created.
3) Copy the Class Activites folder to the new WorkingFolder.

## 4. Create a Personal Access Token in GitHub
Follow these steps:
1) Login to https://github.com/
2) Click the icon in the top right corner of GitHub
3) Choose Settings
4) Scroll down on the left and choose Developer Settings
5) Choose Personal Access Tokens on the left
6) Pick Tokens (Classic) on the left
7) On the right, click Generate New Token
8) Then pick, Generate New Token (classic)
9) For the note, enter - MyGitHubPAT
10) Change the Expiration to Custom
11) Enter 12/31/2024
12) Under Select Scopes put a checkmark next to **repo**
13) Scroll all the way down and click the green Generate Token button
14) Now, you will see your token. It looks like this:
		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
15) Copy that to the clipboard with the copy button
16) Open Notepad (from the start menu)
17) Paste in the token.
18) Save it to your documents folder (NOT one of the folders under DataViz2024)
19) We'll need it later, so hang on to it, and leave that file open

## 5. Clone Simon's repo
Follow these steps:
1) Open a browser to https://github.com/simonkingaby/DataViz2024-Public
2) Click the green Code button
3) Copy the Https URL
4) Back in Git Bash, navigate to the DataViz2024 folder
		You might need to go up a folder, type 
		```cd ..```
5) Make sure you are not in the UNC-VIRT folder
![[Pasted image 20240629110100.png]]
7) Enter the commands
```bash
git clone https://github.com/simonkingaby/DataViz2024-Public.git
cd DataViz2024-Public
```
8) Configure your user in git bash
```bash
git config user.email "your-email@example.com"
git config user.name "Your GitHub User Name" 
# for example, git config user.name "simonkingaby"
git lfs install
# then do a git pull to test it all
git pull
```

## 6. Create public repos for Homework
Follow these steps:
1) Back in https://github.com/
2) On the left, near Top Repositories, click the green New button
3) Skip the template
4) Make sure you are the owner
5) Enter Homework-01 for the Module 01 homework, Homework-02 for the Module 02 homework, and so on.
6) Select Public 
7) Check the Add a README file box
8) In the Add .gitignore select Python
9) Then click the green Create Repository
10) Repeat these steps for each module

## 7. Clone the Homework repos
1) With the repo created click the green Code button
2) Copy the Https URL
Back in Git Bash, navigate to the DataViz2024 folder
		You might need to go up a folder, type 
		```cd ..```
3) Make sure you are not in the UNC-VIRT or DataViz2024-Public folder
![[Pasted image 20240629110100.png]]
4) Enter the commands
```bash
git clone <paste in your URL here>
cd <your new homework folder> 
```
5) Configure your user in git bash
```bash
git config user.email "your-email@example.com"
git config user.name "Your GitHub User Name" 
# for example, git config user.name "simonkingaby"
git lfs install
# then do a git pull to test it all
git pull
```
6) Repeat for each Homework repo
## 8. Celebrate
Take a break to celebrate because that's a job well done.

![homer-computer-woohoo.jpg (788×500) (bobleesays.com)](https://bobleesays.com/wp-content/uploads/2015/05/homer-computer-woohoo.jpg)
