# Developer Pre-requisites

# Set up your Office 365 tenant

To build and deploy PowerApps, Flows, and client-side web parts using the SharePoint Framework, you need an Office 365 tenant. 

If you don't have one, you can get an Office 365 developer subscription when you join the [Office 365 Developer Program](http://aka.ms/M365devbootcamp2019attendees). See the [Office 365 Developer Program documentation](https://docs.microsoft.com/en-us/office/developer-program/office-365-developer-program) for step-by-step instructions about how to join the Office 365 Developer Program and sign up and configure your subscription.  

> [!NOTE] 
> Make sure that you are signed out of any existing Office 365 tenants before you sign up for the Office 365 Developer Program.

# Set up your SharePoint Framework development environment

## Developer tools

## Prerequisites

Developing apps for Microsoft Teams requires preparation for both the Office 365 tenant and the development workstation.

For the Office 365 Tenant, the setup steps are detailed on the [Prepare your Office 365 tenant page](https://docs.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-tenant). Note that while the getting started page indicates that the Public Developer Preview is optional, this lab includes steps that are not possible unless the preview is enabled. Information about the Developer Preview program and participation instructions are detailed on the [What is the Developer Preview for Microsoft Teams? page](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/dev-preview/developer-preview-intro).

### Azure Subscription

The Azure Bot service requires an Azure subscription. A free trial subscription is sufficient.

If you do not wish to use an Azure Subscription, you can use the legacy portal to register a bot here: [Legacy Microsoft Bot Framework portal](https://dev.botframework.com/bots/new) and sign in. The bot registration portal accepts a work, school account or a Microsoft account.

### Install developer tools

The developer workstation requires the following tools for this lab.

#### Install NodeJS & NPM

Install [NodeJS](https://nodejs.org/) Long Term Support (LTS) version. If you have NodeJS already installed please check you have the latest version using `node -v`. It should return the current [LTS version](https://nodejs.org/en/download/). Allowing the **Node setup** program to update the computer `PATH` during setup will make the console-based tasks in this easier to accomplish.

After installing node, make sure **npm** is up to date by running following command:

````shell
npm install -g npm
````

#### Install Yeoman, Gulp-cli and TypeScript

[Yeoman](http://yeoman.io/) helps you start new projects, and prescribes best practices and tools to help you stay productive. This lab uses a Yeoman generator for Microsoft Teams to quickly create a working, JavaScript-based solution. The generated solution uses Gulp, Gulp CLI and TypeScript to run tasks.

Enter the following command to install the prerequisites:

````shell
npm install -g yo gulp-cli typescript
````

#### Install Yeoman Teams generator

The Yeoman Teams generator helps you quickly create a Microsoft Teams solution project with boilerplate code and a project structure & tools to rapidly create and test your app.

Enter the following command to install the Yeoman Teams generator:

````shell
npm install generator-teams -g
````

#### Download ngrok

As Microsoft Teams is an entirely cloud-based product, it requires all services it accesses to be available from the cloud using HTTPS endpoints. To enable the exercises to work within Microsoft Teams, a tunneling application is required.

This lab uses [ngrok](https://ngrok.com) for tunneling publicly-available HTTPS endpoints to a web server running locally on the developer workstation. ngrok is a single-file download that is run from a console.

#### Code editors

Tabs in Microsoft Teams are HTML pages hosted in an iframe. The pages can reference CSS and JavaScript like any web page in a browser.

Microsoft Teams supports much of the common [bot framework](https://dev.botframework.com/) functionality. The Bot Framework provides an SDK for C# and Node.

You can use any code editor or IDE that supports these technologies, however the steps and code samples in this training use [Visual Studio Code](https://code.visualstudio.com/) for tabs using HTML/JavaScript and [Visual Studio 2017](https://www.visualstudio.com/) for bots using the C# SDK.

#### Install [.NET Core SDK](https://dotnet.microsoft.com/download) version 2.1

  ```bash
  # determine dotnet version
  dotnet --version
  ```

#### Bot template for Visual Studio 2017

Download and install the [bot template for C#](https://github.com/Microsoft/BotFramework-Samples/tree/master/docs-samples/CSharp/Simple-LUIS-Notes-Sample/VSIX) from Github. Additional step-by-step information for creating a bot to run locally is available on the [Create a bot with the Bot Builder SDK for .NET page](https://docs.microsoft.com/en-us/azure/bot-service/dotnet/bot-builder-dotnet-quickstart?view=azure-bot-service-3.0) in the Azure Bot Service documentation.

  > **Note:** This lab uses the BotBuilder V3 SDK. BotBuilder V4 SDK was recently released. All new development should be targeting the BotBuilder V4 SDK. In our next release, this sample will be updated to the BotBuilder V4 SDK.

<a name="exercise1"></a>

#### Install Git command line tools
Install the [Git source control tools](https://git-scm.com/)

## Clone our starter repo

### Create a working folder
Create a working folder where you want to work on the bootcamp SharePoint Framework labs. We recommend using a simple folder structure, something like "c:\Work\bootcamp" on Windows, or "~/code/bootcamp" on a Mac.

### Clone the repository from Github
Open a command prompt in your working folder, and then run the following command:
```sh
git clone https://github.com/LK-MUG/2019-Global-M365-Developer-Bootcamp.git
```
