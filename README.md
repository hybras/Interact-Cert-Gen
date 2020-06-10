# Generating Certificates

Install gradle and java(at least 11) on your system. Then just `./gradlew run` and you're good to go. Make sure to convert the certificates to pdfs afterwards.

## Getting Started

1. Install at least java 11
   1. Java is the language this project is written in, gradle is the program that will build and run the project
2. Download the project.
   1. See the big green download button in the upper right. You can either `git clone` it, or download the zip file.
3. Open the command line and navigate to the folder in which you downloaded the project. This step will vary based on what shell and operating system you are using.
4. Set the template and attendance spreadsheet
   1. Go to the templates folder.
   2. Replace the attendance spreadsheet with the latest version.
   3. You may edit the certificate template as well. Follow the syntax for the fill-and-replace. The program looks for the words "NAME," "EVENT," "HOURS", and "DATE". Feel free to modify the program to accept other replacements.Just open a pr with your changes.
5. Run `.\gradlew run`
   1. If this is the first time running the project, it will need to build first. This will take about a minute, then run afterwards.
6. The program will take longer if there are more events or people. It will output to a certificates folder. Do not edit these files till the program is done. You can view the progress by viewing files as they are added to the folder.
