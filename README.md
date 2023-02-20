## Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps:

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

2. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:

   ```console
   npm install --global http-server
   ```

3. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet you can do this with the following command:

   ```console
   npm install --global office-addin-dev-certs
   ```

4. Clone or download this sample to a folder on your computer. Then go to that folder in a console or terminal window.
5. Run the following command to generate a self-signed certificate that you can use for the web server.

   ```console
   npx office-addin-dev-certs install
   ```

   The previous command will display the folder location where it generated the certificate files.

6. Go to the folder location where the certificate files were generated. Copy the localhost.crt and localhost.key files to the hello world sample folder.

7. Run the following command:

   ```console
   sudo npm run dev-server
   ```
   
   
