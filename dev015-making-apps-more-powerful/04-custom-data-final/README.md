# Adding custom data to Microsoft Graph resources

This exercise will walk you through working with custom data for resources using Microsoft Graph.
It uses .Net Core 2.1, MSAL, and raw REST requests to Microsoft Graph.
It also uses the newly (preview) supported device code flow using MSAL.

## Pre-requisites

1. VS 2017 or VS Code
2. .NET Core 2.1 SDK
3. If using VS Code, you'll need the **C# for Visual Studio Code** extension
4. An administrative user

## Instructions

### Getting started

1. Clone or download the repo from here: [https://github.com/Azure-Samples/active-directory-coding-exercises](https://github.com/Azure-Samples/active-directory-coding-exercises)
2. Go to the `dev015-making-apps-more-powerful\04-custom-data` folder and open the solution file in VS 2017, or in VS Code open the `04-custom-data` folder.
    1. When opening in VS Code, click yes to resolve any required C# elements and also click to restore packages.
    2. If package restore doesn't work as part of building the solution, you can do this manually too: Terminal -> New Terminal, and then in the new terminal, type: 
        * `dotnet add package Microsoft.Graph`, and hit Restore button when prompted (if prompted)
        * `dotnet add package Microsoft.Identity.Client`, and hit Restore button when prompted (if prompted)
3. Go to the private preview app registration experience at [http://aka.ms/registeredappsprod](http://aka.ms/registeredappsprod). 
    1. Register a new single tenant **public client** app (on the first creation page, change "web app" to "public client" and create).  
    2. Go to Manage -> Manifest and open the app manifest. (The app manifest is simply a GET/PATCH on the `/application` resource in Microsoft Graph.) Due to a bug, you'll need to change the `allowPublicClient` property to `true`.
    3. Now get the client id. You'll need this for the next step.
4. Update the `appsettings.json` file with your app's co-ordinates (clientId and tenant - like contoso.onmicrosoft.com).
5. In VS Code, Debug -> Open Configuration, and then click "Add Configuration" button. Then select an appropriate option (.NET: Launch .Net Core Console App)
    * Change "program" to the same execuatable name as in the existing .Net Core console app config
    * Change "console" to "externalTerminal" (not sure if this is platform independent - may need to look online)
    * Change "name" to "External Console"
  
6. Start Debugging (or F5/equivalent), but switch debug to use "External Console"
7. In the partially built state, you'll be asked to go through the device code flow.
    1.  Per the instructions in the console, copy the `code`, and go to [https://mmicrosoft.com/devicelogin](https://mmicrosoft.com/devicelogin).
    2.  Enter the code and click continue
    3.  Now follow the standard flow for sign in

    You should now be signed in to the app, and in the console window you'll see the signed-in user's profile and manager, open extensions and schema extensions manipulations.