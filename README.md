# Simple Outlook-Addin

This Contains slightly modified and simpler version of Addin which has only taskpane.

[Original Source](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/hello-world/outlook-hello-world)

## Summary

Learn how to build the simplest Office Add-in with only a manifest, HTML web page, and a logo. This sample will help you understand the fundamental parts of an Office Add-in.

### Manifest

The manifest file is an XML file that describes your add-in to Office. It contains information such as a unique identifier, name, what buttons to show on the ribbon, and more. Importantly the manifest provides URL locations for where Office can find and download the add-in's resource files.

## Run the sample

An Office Add-in requires you to configure a web server to provide all the resources, such as HTML, image, and JavaScript files. Select one of the following options to run the hello world sample.

### Run the sample using GitHub as a web host

The hello world sample is configured so that the files are hosted directly from this GitHub repo.

1. Download the **manifest.xml** file from this sample to a folder on your computer.
1. Sideload the manifest in Outlook on Windows (new or classic), on Mac, or on the web by following the instructions in [Sideload Outlook add-in on Windows or Mac](https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing).

### Configure a localhost web server and run the sample from localhost

If you prefer to configure a web server and host the add-in's web files from your computer, use the following steps.

1. Install a recent version of [npm](https://www.npmjs.com/get-npm) and [Node.js](https://nodejs.org/) on your computer. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.

1. You need http-server to run the local web server. If you haven't installed this yet you can do this with the following command:

   ```console
   npm install --global http-server
   ```

1. You need Office-Addin-dev-certs to generate self-signed certificates to run the local web server. If you haven't installed this yet you can do this with the following command:

   ```console
   npm install --global office-addin-dev-certs
   ```

1. Clone or download this sample to a folder on your computer. Then, go to that folder in a console or terminal window.
1. Run the following command to generate a self-signed certificate to use for the web server.

   ```console
   npx office-addin-dev-certs install
   ```

   This command will display the folder location where it generated the certificate files.

1. Go to the folder location where the certificate files were generated. Copy the **localhost.crt** and **localhost.key** files to the cloned or downloaded sample folder.

1. Run the following command.

   ```console
   http-server -S -C localhost.crt -K localhost.key --cors . -p 3000 -a localhost
   ```

   If you don't use `-a localhost` nothing much happens only it will run on `127.0.0.1`.
   The http-server will run and host the current folder's files on `localhost:3000`.
