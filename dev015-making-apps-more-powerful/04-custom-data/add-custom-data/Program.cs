﻿/*
 The MIT License (MIT)

Copyright (c) 2015 Microsoft Corporation

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace add_custom_data
{
    /// <summary>
    /// This sample signs-in a user in a two steps process:
    /// - it displays a URL and a code, and asks the user to navigate to the URL in a Web browser, and enter the code
    /// - then the user signs-in (and goes through multiple factor authentication if needed)
    /// and the sample displays information about the user by calling the Microsoft Graph in the name of the signed-in user
    /// 
    /// It uses the Device code flow, which is normally used for devices which don't have a Web browser (which is the case for a
    /// .NET Core app, iOT, etc ...)
    /// 
    /// For more information see https://aka.ms/msal-net-device-code-flow
    /// </summary>
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                RunAsync(args).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

        }

        static async Task RunAsync(string[] args)
        {
            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");
            var app = new PublicClientApplication(config.ClientId, config.Authority);

            MyInformation myInformation = new MyInformation(app);
            await myInformation.DisplayMeAndMyManagerAsync();

            var openExtensionsDemo = new OpenExtensionsDemo(app);
            await openExtensionsDemo.RunAsync();

           var schemaExtensionDemo = new SchemaExtensionsDemo(app);
           await schemaExtensionDemo.RunAsync();

            System.Console.WriteLine("Press ENTER to continue.");
            System.Console.ReadLine();
        }

    }
}
