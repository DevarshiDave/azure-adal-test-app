# azure-adal-test-app
This is an Office Outlook Addin created using **yo office**.

To start Project use **npm install** command first, then use **npm start**.

Also you need to register an app in your azure active directory with below permissions
  1. Sharepoint site collection read/write 
  2. Microsoft Groups read/write of microsoft graph api.

Set up your azure app to **allow Implicit flow authentication** by setting its value to **true** in the manifest file.

Copy the **Application ID** given in Azure Active directory App and paste it in **clientid** object given in **app.js**

and you're good to go.

Cheers!
