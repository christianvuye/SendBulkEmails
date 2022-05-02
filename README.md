# SendBulkEmails

- **Description:** This project is a small desktop application I wrote for a friend, who needed an easy tool to send email with the same template but different recipient names in bulk to different email addresses. The formatting of the output is hard-coded for the specific template of said friend, but can easily be changed by adjusting the _"css"_ variable to whatever CSS the user prefers. The script uploaded on this repository handles one specific template, but multiple templates can be used in the Excel sheet. With a small tweak to the _for loop_ in the script logic, the script is also capable of handling multiple templates. In the future, this functionality might be added if said friend, me or someone else needs it. 

-  **Features:** This application reads email address, name, subject and template input from an _.xlsx sheet_, along with text from a _.docx_  template. It converts the text to _HTML_ in order to make sure other elements such as tables are included in the email that is sent. The script also adds a small bit of custom _CSS_ in order to make sure the converted HTML is styled correctly.

-  **How to use:** The script requires specific user security-related _.json_ files to be present in the same folder as the script or executable. Finding where to download these files from can be quite challenging if you don't know where to look. As an elaborate tutorial on how to download these files is outside of the scope of this document, a _"Tutorial"_ folder will be uploaded with both a written as a video walk-through of how to set the application up.
 
-  **Technologies:** 
		+ Python
		+ HTML
		+ CSS
		+ ezgmail  
		+ mammoth
		+ openpyxl

-  **Collaborators:** Sebastian Vuye (https://github.com/sebavuye) provided support on HTML and CSS as well as being a rubber duck for when I ran into issues.
-   **License:** MIT License Copyright (c) 2022 Christian Vuye