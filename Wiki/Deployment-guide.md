- [Deployment Guide](#outlook-web-service-ows)
    - [Prerequisites](#prerequisites)
    - [Steps](#Deployment-Steps)
        - [Register AD Application](#1-register-azure-ad-application)
        - [Deploy to Azure subscription](#2-deploy-to-your-azure-subscription)
        - [Set-up Authentication](#3-set-up-authentication)
        - [Add Permissions to your app](#4-add-permissions-to-your-app)
        - [Create the Teams app packages](#5-create-the-teams-app-packages)
        - [Install the apps in Microsoft Teams](#6-install-the-apps-in-microsoft-teams)
    - [Troubleshooting](#troubleshooting)
- - -

# Prerequisites
To begin, you will need: 
* An Azure subscription where you can create the following kinds of resources:  
    * App Service
    * App Service Plan
    * Bot Channels Registration
    * Azure Function
    * Azure Storage Account
    * Service Bus
    * Application Insights
* A team with the users who will be sending messages with this app. (You can add or remove team members later!)
* A copy of the Company Communicator app GitHub repo (https://github.com/OfficeDev/microsoft-teams-company-communicator-app)

> **NOTE:** If you plan to use a custom domain name instead of relying on Azure Front Door, read the instructions [here](Custom-domain-option) first.

- - -

# Deployment Steps

## 1. Register Azure AD application

Register two Azure AD application in your tenant's directory: one for author bot, and another for user bot.

1. Log in to the Azure Portal for your subscription, and go to the [App registrations](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps) blade.

1. Click **New registration** to create an Azure AD application.
    - **Name**: Name of your Teams App - if you are following the template for a default deployment, we recommend "Company Communicator".
    - **Supported account types**: Select "Accounts in any organizational directory" (*refer image below*).
    - Leave the "Redirect URI" field blank for now.

    ![Azure AD app registration page](images/multitenant_app_creation.png)

1. Click **Register** to complete the registration.

1. When the app is registered, you'll be taken to the app's "Overview" page. Copy the **Application (client) ID**; we will need it later. Verify that the "Supported account types" is set to **Multiple organizations**.

    ![Azure AD app overview page](images/multitenant_app_overview_1.png)

1. On the side rail in the Manage section, navigate to the "Certificates & secrets" section. In the Client secrets section, click on "+ New client secret". Add a description for the secret, and choose when the secret will expire. Click "Add".

    ![Azure AD app secret](images/multitenant_app_secret.png)

1. Once the client secret is created, copy its **Value**; we will need it later.

1. Go back to "App registrations", then repeat steps 2-5 to create another Azure AD application for the author bot.
    - **Name**: Name of your Teams App - if you are following the template for a default deployment, we recommend "Company Communicator Author".
    - **Supported account types**: Select "Accounts in any organizational directory".
    - Leave the "Redirect URI" field blank for now.


    At this point you should have the following 5 values:
    1. Application (client) ID for the user bot.
    2. Client secret for the user bot.
    3. Directory (tenant) ID.
    4. Application (client) Id for the author bot.
    5. Client secret for the author bot.

    We recommend that you copy the values, we will need them later.

    ![Azure AD app overview page](images/multitenant_app_overview_2.png)

## 2. Deploy to your Azure subscription
1. Click on the **Deploy to Azure** button below.
   
   [![Deploy to Azure](https://azuredeploy.net/deploybutton.png)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2Fcristianoag%2Fmicrosoft-teams-company-communicator-app%2Fmaster%2FDeployment%2Fazuredeploy.json)

1. When prompted, log in to your Azure subscription.

1. Azure will create a "Custom deployment" based on the Company Communicator ARM template and ask you to fill in the template parameters.

    > **Note:** Please ensure that you don't use underscore (_) or space in any of the field values otherwise the deployment may fail.

1. Select a subscription and a resource group.
   * We recommend creating a new resource group.
   * The resource group location MUST be in a datacenter that supports all the following:
     * Storage Accounts
     * Application Insights
     * Azure Functions
     * Service Bus
     * App Service

    For an up-to-date list of datacenters that support the above, click [here](https://azure.microsoft.com/en-us/global-infrastructure/services/?products=storage,app-service,monitor,service-bus,functions)

1. Enter a **Base Resource Name**, which the template uses to generate names for the other resources.
   * The `[Base Resource Name]` must be available. For example, if you select `contosocommunicator` as the base name, the name `contosocommunicator` must be available (not taken); otherwise, the deployment will fail with a Conflict error.
   * Remember the base resource name that you selected. We will need it later.

1. Update the following fields in the template:
    1. **User Client ID**: The application (client) ID of the Microsoft Teams user bot app. (from Step 1)
    2. **User Client Secret**: The client secret of the Microsoft Teams user bot app. (from Step 1)
    3. **Tenant Id**: The tenant ID. (from Step 1)
    4. **Author Client ID**: The application (client) ID of the Microsoft Teams author bot app. (from Step 1)
    5. **Author Client Secret**: The client secret of the Microsoft Teams author bot app. (from Step 1)
    6. **Proactively Install User App [Optional]**: Default value is `true`. You may set it to `false` if you want to disable the feature.
    7. **User App ExternalId [Optional]**: Default value is `148a66bb-e83d-425a-927d-09f4299a9274`. This **MUST** be the same `id` that is in the Teams app manifest for the user app.
    8. **DefaultCulture [Optional]**: By default the application uses `en-US` locale. You can choose the locale from the list, if you wish to use the app in different locale.Also, you may add/update the resources for other locales and update this configuration if desired.
    9. **SupportedCultures [Optional]**: This is the list of locales that application supports currently.You may add/update the resources for other locales and update this configuration if desired.

    > **Note:** Make sure that the values are copied as-is, with no extra spaces. The template checks that GUIDs are exactly 36 characters.

    > **Note:** If your Azure subscription is in a different tenant than the tenant where you want to install the Teams App, please update the `Tenant Id` field with the tenant where you want to install the Teams App.

1. Update the "Sender UPN List", which is a semicolon-delimited list of users (Authors) who will be allowed to send messages using the Company Communicator.
    * For example, to allow Megan Bowen (meganb@contoso.com) and Adele Vance (adelev@contoso.com) to send messages, set this parameter to `meganb@contoso.com;adelev@contoso.com`.
    * You can change this list later by going to the App Service's "Configuration" blade.

1. If you wish to change the app name, description, and icon from the defaults, modify the corresponding template parameters.

1. Agree to the Azure terms and conditions by clicking on the check box "I agree to the terms and conditions stated above" located at the bottom of the page.

1. Click on "Purchase" to start the deployment.

1. Wait for the deployment to finish. You can check the progress of the deployment from the "Notifications" pane of the Azure Portal. It may take **up to an hour** for the deployment to finish.

    > If the deployment fails, see [this section](https://github.com/OfficeDev/microsoft-teams-company-communicator-app/wiki/Troubleshooting#1-code-deployment-failure) of the Troubleshooting guide.

1. Once the deployment is successfully completed, go to the deployment's "Outputs" tab, and note down the follwing values. We will need them later.
    * **authorBotId:** This is the Microsoft Application ID for the Company Communicator app. For the following steps, it will be referred to as `%authorBotId%`.
    * **userBotId:** This is the Microsoft Application ID for the Company Communicator app. For the following steps, it will be referred to as `%userBotId%`.
    * **appDomain:** This is the base domain for the Company Communicator app. For the following steps, it will be referred to as `%appDomain%`.

> **IMPORTANT:** If you plan to use a custom domain name instead of relying on Azure Front Door, read the instructions [here](Custom-domain-option) before continuing any further.

## 3. Set-up Authentication

1. Note that you have the `%authorBotId%`, `%userBotId%` and `%appDomain%` values from the previous step (Step 2).

    > If do not have these values, refer [this section](https://github.com/OfficeDev/microsoft-teams-company-communicator-app/wiki/Troubleshooting#2-forgetting-the-botId-or-appDomain) of the Troubleshooting guide for steps to get these values.

1. Go to **App Registrations** page [here](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps) and open the author app you created (in Step 1) from the application list.

1. Under **Manage**, click on **Authentication** to bring up authentication settings.

    1. Add a new entry to **Redirect URIs**:
        - **Type**: Web
        - **Redirect URI**: Enter `https://%appDomain%/signin-simple-end` for the URL e.g. `https://appName.azurefd.net/signin-simple-end`

    1. Under **Implicit grant**, check **ID tokens**.

    1. Click **Save** to commit your changes.

1. Back under **Manage**, click on **Expose an API**.

    1. Click on the **Set** link next to **Application ID URI**, and change the value to `api://%appDomain%` e.g. `api://appName.azurefd.net`.

    1. Click **Save** to commit your changes.

    1. Click on **Add a scope**, under **Scopes defined by this API**. In the flyout that appears, enter the following values:
        * **Scope name:** access_as_user
        * **Who can consent?:** Admins and users
        * **Admin and user consent display name:** Access the API as the current logged-in user
        * **Admin and user consent description:**  Access the API as the current logged-in user

    1. Click **Add scope** to commit your changes.

    1. Click **Add a client application**, under **Authorized client applications**. In the flyout that appears, enter the following values:
        * **Client ID**: `5e3ce6c0-2b1f-4285-8d4b-75ee78787346`
        * **Authorized scopes**: Select the scope that ends with `access_as_user`. (There should only be 1 scope in this list.)

    1. Click **Add application** to commit your changes.

    1. **Repeat the previous two steps**, but with client ID = `1fec8e78-bce4-4aaf-ab1b-5451cc387264`. After this step you should have **two** client applications (`5e3ce6c0-2b1f-4285-8d4b-75ee78787346` and `1fec8e78-bce4-4aaf-ab1b-5451cc387264`) listed under **Authorized client applications**.

1. Back under **Manage**, click on **Manifest**.

   1. In the editor that appears, find the `optionalClaims` property in the JSON Azure AD application manifest, and replace it with the following block:
    ```
        "optionalClaims": {
            "idToken": [],
            "accessToken": [
                {
                    "name": "upn",
                    "source": null,
                    "essential": false,
                    "additionalProperties": []
                }
            ],
            "saml2Token": []
        },
    ```

    2. Click **Save** to commit your changes.

## 4. Add Permissions to your app

Continuing from the Azure AD app registration page where we ended Step 3.

1. Select **API Permissions** blade from the left hand side.

2. Click on **Add a permission** button to add permission to your app.

3. In Microsoft APIs under Select an API label, select the particular service and give the following permissions,

    * Under **Commonly used Microsoft APIs**, 

    * Select “Microsoft Graph”, then select **Delegated permissions** and check the following permissions,
        1. **Group.Read.All**
        2. **AppCatalog.Read.All**

    * then select **Application permissions** and check the following permissions,
        1. **Group.Read.All**
        2. **User.Read.All**
        3. **TeamsAppInstallation.ReadWriteForUser.All**

    * Click on **Add Permissions** to commit your changes.

    ![Azure AD API permissions](images/multitenant_app_permissions_1.png)
    ![Azure AD API permissions](images/multitenant_app_permissions_2.png)

    > Please refer to [Solution overview](https://github.com/OfficeDev/microsoft-teams-company-communicator-app/wiki/Solution-overview#microsoft-graph-api) for more details about the above permissions.

4. If you are logged in as the Global Administrator, click on the “Grant admin consent for %tenant-name%” button to grant admin consent, else inform your Admin to do the same through the portal.
   <br/>
   Alternatively you may follow the steps below:
   - Prepare link - https://login.microsoftonline.com/common/adminconsent?client_id=%appId%. Replace the `%appId%` with the `Application (client) ID` of Microsoft Teams author bot app (from above).
   - Global Administrator can grant consent using the link above.

## 5. Create the Teams app packages

Company communicator app comes with 2 applications – Author, User. The Author application is intended for employees who create and send messages in the organization, and the User application is intended for employees who receive the messages.

Create two Teams app packages: one to be installed to an Authors team and other for recipients to install personally and/or to teams.

1. Make sure you have cloned the app repository locally.

1. Open the `Manifest\manifest_authors.json` file in a text editor.

1. Change the placeholder fields in the manifest to values appropriate for your organization.
    * `developer.name` ([What's this?](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema#developer))
    * `developer.websiteUrl`
    * `developer.privacyUrl`
    * `developer.termsOfUseUrl`

1. Change the `<<appDomain>>` placholder in the configurationUrl setting to be the `%appDomain%` value e.g. "`https://appName.azurefd.net/configtab`".

1. Change the `<<botId>>` placeholder in the botId setting to be the `%authorBotId%` value - this is your author Azure AD application's ID from above. This is the same GUID that you entered in the template under "Author Client ID". Please note that there are two places in the manifest (for authors) where you will need to update Bot ID.

1. Change the `<<appDomain>>` placeholder in the validDomains setting to be the `%appDomain%` value e.g. "`appName.azurefd.net`".

1. Change the `<<botId>>` placeholder in the id setting of the webApplicationInfo section to be the `%authorBotId%` value. Change the `<<appDomain>>` placeholder in the resource setting of the webApplicationInfo section to be the `%appDomain%` value e.g. "`api://appName.azurefd.net`".

1. Copy the `manifest_authors.json` file to a file named `manifest.json`.

1. Create a ZIP package with the `manifest.json`,`color.png`, and `outline.png`. The two image files are the icons for your app in Teams.
    * Name this package `company-communicator-authors.zip`, so you know that this is the app for the author teams.
    * Make sure that the 3 files are the _top level_ of the ZIP package, with no nested folders.  
    ![image10](images/file-explorer.png)

1. Delete the `manifest.json` file.

Repeat the steps above but with the file `Manifest\manifest_users.json` and use `%userBotId%` for `<<botId>>` placeholder. Note: you will not need to change anything for the configurationUrl or webApplicationInfo section because the recipients app does not have the configurable tab. Name the resulting package `company-communicator-users.zip`, so you know that this is the app for the recipients.

## 6. Install the apps in Microsoft Teams

1. Install the authors app (the `company-communicator-authors.zip` package) to your team of message authors.
    * Note that even if non-authors install the app, the UPN list in the app configuration will prevent them from accessing the message authoring experience. Only the users in the sender UPN list will be able to compose and send messages. 
    * If your tenant has sideloading apps enabled, you can install your app by following the instructions [here](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload#load-your-package-into-teams).

2. Add the configurable tab to the team of authors, so that they can compose and send messages.

3. [Upload](https://docs.microsoft.com/en-us/microsoftteams/tenant-apps-catalog-teams) the User app to your tenant's app catalog so that it is available for everyone in your tenant to install.
> **IMPORTANT:** Proactive app installation will work only if you upload the User app to your tenant's app catalog.

4. Install the User app (the `company-communicator-users.zip` package) to the users and teams that will be the target audience.
> If `proactiveAppInstallation` is enabled, you may skip this step. The service will install the app for all the recipients when authors send a message.

> **NOTE:** If you are deploying a version of Company Communicator prior to version 4, do NOT use app permission policies to restrict the authors app to the members of the authors team. Microsoft Teams does not support applying different policies to the same bot via two different app packages. 

---

# Troubleshooting
Please check the [Troubleshooting](Troubleshooting) guide.
