# Sherpa/Romeo Script
## Description:
The project uses a Google Apps Script that utilizes the Sherpa/RoMEO API to retrieve journal permissions data into a Google sheet. The goal of the script is to search batches of journal ISSNs and identify permissions, conditions, and restrictions for submitting published scholarly research to an institutional repository.

1. Register for a Sherpa/RoMEO API Key: http://www.sherpa.ac.uk/romeo/apiregistry.php 
2. Create a Google sheet, click on the Tools menu, and select "Script editor." 
3. Copy and paste the contents of SherpaRoMEO_Permissions.gs (be sure to replace the existing default code). 
4. Enter your Sherpa API Key between the quotes.
5. Click File -> Save, enter the project name "Sherpa/RoMEO Permissions" and close the Script editor browser tab. 
6. Click the browser refresh button to add the SherpaRoMEO menu to your Google sheet.
7. Click on SherpaRoMEO and select "Create Column Headers." 
8. Authroize Script
   1. If prompted to authorize, click continue. Then, click on your Google account.
   2. If prompted that the app isn't verified, click Advanced, scroll down, and click on "Go to Sherpa/RoMEO Permissions (unsafe)"
   3. If prompted to give the project access to your Google account, scroll down and click "Allow"
9. Add ISSNs to the ISSN column. Click on SherpaRoMEO and select "Process ISSNs."

Step 8 is only required once during the initial setup.

## License
The content of this project is licensed under the [Creative Commons Attribution 4.0 license](https://creativecommons.org/licenses/by/4.0/).
 
