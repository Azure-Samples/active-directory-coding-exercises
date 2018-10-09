# User changes delta query exercise

This is a .NET Core console app and can be loaded, run and debugged through VS 2017 or VS Code.  It uses
both the Microsoft Authentication Library (MSAL) and the Microsoft Graph SDK.
It's partially completed, and the instuctions walk through the steps you'll need
to take to complete the exercise - by adding code and guidance on what to add in the right places, such as:

```C#
/// <exercise_hint>
/// Some hints that can help you.
/// This topic also provides relevant background https://developer.microsoft.com/en-us/graph/docs/concepts/delta_query_users
/// </exercise_hint>

/// ADD CODE HERE
```

All changes in this exercise are in the `Program.cs` file, although the getting started steps include some changes to the configuration in `Config.cs`

## Pre-requisites

1. VS 2017 or VS Code
2. .NET Core 2.1 SDK
3. If using VS Code, you'll need the **C# for Visual Studio Code** extension

## Instructions

### Getting started

1. Clone or download the repo from here: [https://github.com/Azure-Samples/active-directory-coding-exercises](https://github.com/Azure-Samples/active-directory-coding-exercises)
2. Go to the `dev015-making-apps-more-powerful\03-batch` folder and open the solution file in VS 2017, or in VS Code open the `03-batch` folder.
    1. When opening in VS Code, click yes to resolve any required C# elements and also click to restore packages.
    2. If package restore doesn't work as part of building the solution, you can do this manually too: Terminal -> New Terminal, and then in the new terminal, type: 
        * `dotnet add package Microsoft.Graph`, and hit Restore button when prompted (if prompted)
        * `dotnet add package Microsoft.Identity.Client`, and hit Restore button when prompted (if prompted)
1. Go to the private preview app registration experience at [http://aka.ms/appregprivatepreview](http://aka.ms/appregprivatepreview) . Register a new single tenant app, and create a new secret.  Also get the client id. Configure the app with **User.Read.All** and **Sites.Read.All** permissions and then **grant** the app this permission.
3. Update the Config class in the Program.cs file with your app's co-ordinates.
4. In VS Code, Debug -> Open Configuration, and then click "Add Configuration" button. Then select an appropriate option (.NET: Launch .Net Core Console App)
  * Change "program" to the same execuatable name as in the existing .Net Core console app config
  * Change "console" to "externalTerminal" (not sure if this is platform independent - may need to look online)
  * Change "name" to "External Console"
8. Start Debugging (or F5/equivalent), but switch debug to use "External Console"
9. In the partially built state, the app should run, and appear to create a user, ask you to press a key, and then it cleans up, and is done.

### Excercise 

Generate the batch request in order to fetch below data

1. User's first and last name along with job title.
2. Recent documents that user has been working on.

Ensure documents are not fetched in case of error accessing the user profile.