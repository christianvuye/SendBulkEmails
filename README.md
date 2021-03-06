# SendBulkEmails

- **Description:** This project is a small desktop application I wrote for a friend who needed an easy tool to send emails with the same template but different recipient names in bulk to separate email addresses. The output formatting is hard-coded for the specific template of the said friend but can change by adjusting the _"CSS"_ variable to whatever CSS the user prefers. The script uploaded to this repository handles one specific template, but the user can add different templates in the Excel sheet. The code can take multiple templates with a slight tweak to the _for loop_ in the logic. I might add this functionality if said friend, I, or someone else needs it in the future. 

- **Features:** This application reads email address, name, subject and template input from a _.xlsx sheet_, along with text from a _.docx_ template. It converts the text to _HTML_ to ensure other elements, such as tables, are included in the sent email. The script also adds a tiny bit of custom _CSS_ to ensure the converted HTML styles are correct.

- **How to use:** The script requires specific user security-related _.json_ files to be present in the same folder as the script or executable. Where to download these files? An elaborate tutorial on downloading them is outside of the scope of this _README_ document. I will upload a _"tutorial"_ folder with a video and written walk-through of how to set the application up. I tested this application with Gmail/Google Workspace only.

- **Technologies:** + Python + HTML + CSS + ezgmail + mammoth + openpyxl

- **Collaborators:** Sebastian Vuye (https://github.com/sebavuye) provided support on HTML and CSS and was a rubber duck for when I ran into issues.

- **License:** MIT License Copyright (c) 2022 Christian Vuye
