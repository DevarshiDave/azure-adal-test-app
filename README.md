# azure-adal-test-app
This is an Office Outlook Addin created using "yo office".

To start Project use "npm install" command first, then use "npm start".

Also you need to have an existing/register app in your azure active directory with Sharepoint site collection read/write permissions and Microsoft Groups read/write permission of microsoft graph api.

Set up your azure app to use Implicit flow authentication.

Copy the Application ID given in Azure Active directory App and paste it in clientid object given in app.js

and you're good to go.

Cheers!
