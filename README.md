# Microsoft Office Add-in Dev Kit for Visual Studio Code
## What is Office Add-in Dev Kit?
The Office Add-in Dev Kit helps developers set up the environment, create and build Office JS add-ins with a steamlined developer experience.

<img src=./assets/Office_Addin_Dev_Kit.gif>

## Getting started
Open Office Add-in Dev Kit to create a new app and start coding!

<img src=./assets/Dev_Kit_GetStarted.png>

In the Office Add-in Dev Kit for Visual Studio Code, you can easily discover all applicable commands in the sidebar and Command Palette with the keyword "Office". It also supports Command Line Interface (CLI) to increase efficiency.

### Prerequisites
Verify you have the right prerequisites for building Office add-ins and install some recommended development tools. [Read more details](https://learn.microsoft.com/office/dev/add-ins/overview/set-up-your-dev-environment).

- **[Node.js](https://nodejs.org) 16, 18, or 20 (18 is preferred) and [npm](https://www.npmjs.com/get-npm).** To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- **Office connected to a Microsoft 365 subscription.** You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](
https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](
https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details.
Alternatively, you can [sign up for a 1-month free trial](
https://www.microsoft.com/microsoft-365/try?rtc=1)
or [purchase a Microsoft 365 plan](
https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).


### Create your add-in
Use the Office Add-in Dev Kit for Visual Studio Code to set up your first add-in project. Create your Office add-in project using the following steps:

* Ensure you've installed the `Microsoft Office Add-in Dev Kit for Visual Studio Code`.
* Select the `Office Add-in Dev Kit` icon in the Visual Studio Code sidebar.
* Select `Create a New Add-in` button from the sidebar.
* Select an Office app that you want to build the add-in for.
* Select an add-in template from the list of available templates.
* Select JavaScript/TypeScript as the programming language.
* Choose a location where your new add-in will be created in a new folder.
* Type a name for your project and hit Enter.

### Configure your add-in

An Office add-in includes two basic components: an XML manifest file, and your own web application. 
* The manifest defines various settings, including how your add-in integrates with Office clients. 
* Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.

When a add-in project is created, you can explore and customize the components by reviewing the key files listed below.

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.  <br>You can check whether your manifest file is valid by selecting `Validate Manifest File` in the `Office Add-in Dev Kit` extension tree view.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/taskpane.css` file contains the CSS that's applied to content in the task pane.
- The `./src/taskpane/taskpane.js` file contains the Office JavaScript API code that facilitates interaction between the task pane and the Word application.

### Preview your add-in in Office apps

To understand how the add-in will work in Office apps, you can use the Office Add-in Dev Kit to easily run and debug your Office add-in in your local dev environment.
Dev_Kit_Preview
<br><img src=./assets/Dev_Kit_Preview.gif>

#### Check and Install Dependencies
Select `Check and Install Dependencies` to check your environment and install necessary dependencies in order to run and debug the add-in code.

#### Preview Your Office Add-in (F5)

Select `Preview Your Office Add-in(F5)` on the side panel to start running and debugging the add-in code. A Word/Excel/PowerPoint app will launch with the add-in sample side-loaded.
- You can also start debugging by hitting the `F5` key or running `npm run start` command in the terminal.
- To debug on Office on the web, go to [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing)
- To debug in Desktop (Edge Legacy), go to [Debug Edge Legacy Webview](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy)

#### Stop Previewing Your Office Add-in

Select `Stop Previewing Your Office Add-in` to stop debugging.


## Explore Code Samples
Explore our [samples](https://github.com/OfficeDev/Office-Samples) to help you quickly get started with the basic Teams app concepts and code structures.

## Reporting security issues
Give security researchers information on how to privately report security vulnerabilities found in your open-source project. See more details [Reporting security issues](https://docs.opensource.microsoft.com/content/releasing/security.html).

## Telemetry
The software may collect information about you and your use of the software and send it to Microsoft. Microsoft may use this information to provide services and improve our products and services. You may turn off the telemetry as described in the repository. There are also some features in the software that may enable you and Microsoft to collect data from users of your applications. If you use these features, you must comply with applicable law, including providing appropriate notices to users of your applications together with a copy of Microsoft's privacy statement. Our privacy statement is located at [Microsoft Privacy Statement](https://privacy.microsoft.com/privacystatement). You can learn more about data collection and use in the help documentation and our privacy statement. Your use of the software operates as your consent to these practices.

## Telemetry Configuration
Telemetry collection is on by default. To opt out, please set the `telemetry.enableTelemetry` setting to `false`. Learn more in our [FAQ](https://code.visualstudio.com/docs/supporting/faq#_how-to-disable-telemetry-reporting).

## Trademark
This project may contain trademarks or logos for projects, products, or services. Authorized use of Microsoft trademarks or logos is subject to and must follow [Microsoft's Trademark & Brand Guidelines](https://www.microsoft.com/legal/intellectualproperty/trademarks/usage/general). Use of Microsoft trademarks or logos in modified versions of this project must not cause confusion or imply Microsoft sponsorship. Any use of third-party trademarks or logos are subject to those third-party's policies.

## License
Copyright (c) Microsoft Corporation. All rights reserved.