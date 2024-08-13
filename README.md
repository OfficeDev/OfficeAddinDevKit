# Microsoft Office Add-ins Development Kit for Visual Studio Code
## What is Office Add-ins Development Kit?
The Office Add-ins Development Kit helps developers set up their environment, create Office Add-ins, and debug their code with a streamlined developer experience.


<img src=./assets/Office_Addin_Dev_Kit.gif alt="The process of using the Office Add-ins Development Kit to create a new add-in. The animation shows a cursor selecting the 'Create a New add-in button', then choosing options for an Excel add-in with a task pane, written in JavaScript."/>


## Getting started
Open Office Add-ins Development Kit to create a new add-in and start coding!

<img src=./assets/Dev_Kit_GetStarted.png alt="Screenshot shows Create a new add-in in Office Add-ins Development Kit"/>

In the Office Add-ins Development Kit for Visual Studio Code, you can discover all the commands in the sidebar and Command Palette with the keyword "Office". It also supports Command Line Interface (CLI) to increase efficiency.


### Prerequisites
Verify you have the right prerequisites for building Office add-ins and install some recommended development tools. [Read more details](https://learn.microsoft.com/office/dev/add-ins/overview/set-up-your-dev-environment).

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify if you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).

### Create your add-in
Use the Office Add-ins Development Kit for Visual Studio Code to set up your first add-in project. Create your Office add-in project using the following steps:

* Ensure you've installed the Microsoft Office Add-ins Development Kit for Visual Studio Code.
* Select the Office Add-ins Development Kit icon in the Visual Studio Code **Activity Bar** to open the extension.
* Select **Create a New Add-in**.
* In the drop down that opens, select an Office app that you want to build the add-in for.
* Select an add-in template from the list of available templates.
* Select JavaScript/TypeScript as the programming language.
* In the **Workspace folder** dialog that opens, select the folder where you want to create the project.
* Give a name to the project (with no spaces) when prompted. Office Add-ins Development Kit will create the project with basic files and scaffolding. It will then open the project *in a second Visual Studio Code window*. Close the original Visual Studio Code window.

### Configure your add-in

An Office add-in includes two basic components: an XML manifest file, and your own web application. 
* The manifest defines various settings, including how your add-in integrates with Office clients. 
* Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.

When a add-in project is created, you can explore and customize the components by reviewing the key files listed below.

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.  <br>You can check whether your manifest file is valid by selecting **Validate Manifest File** option from the Office Add-ins Development Kit.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/taskpane.css` file contains the CSS that's applied to content in the task pane.
- The `./src/taskpane/taskpane.js` file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office application.

### Preview your add-in in Office apps

To understand how the add-in will work in Office apps, you can use the Office Add-ins Development Kit to easily run and debug your Office add-in in your local dev environment.
<br><img src=./assets/Dev_Kit_Preview.gif alt="A user selecting the 'Preview' option from the Office Add-ins Development Kit task pane, then choosing the 'Edge Desktop' option to run the sample."/>


#### Preview Your Office Add-in (F5)

Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the drop down menu, select the option **Desktop (Edge Chromium)**.

The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Office desktop application launches and sideloads the add-in.

You can also start debugging by pressing the `F5` key or running the `npm run start` command in the terminal.

For information on running the add-in on Office on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

#### Stop Previewing Your Office Add-in

Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## See also
All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

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
