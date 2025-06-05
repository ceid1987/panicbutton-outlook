
# Panic Button (Outlook)

This project is a sample add-in for MS Outlook clients that aims to implement a configurable Panic Button to report suspicious emails. 

## Setup

### 1. Environment setup

1. Install [Node.js(LTS)](https://nodejs.org/en/download)

2. Install [git](https://git-scm.com/downloads) and clone the project:

      ```bash
      git clone https://github.com/ceid1987/panicbutton-outlook.git
      ```

3. Inside your project directory, run the following command to install the project dependencies:

      ```bash
      npm install
      ```

### 2. Azure setup

Panic Button uses [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/use-the-api) to run its main functionality. In order to make requests to Graph API, your app/project needs to be registered on [Azure](https://portal.azure.com/).

This step is necessary because as of March 2024, Microsoft has decommissioned the Outlook REST API endpoints.

1. Install [Azure CLI](https://learn.microsoft.com/en-us/cli/azure/install-azure-cli)

2. Install [Yeoman Office Generator](https://github.com/OfficeDev/generator-office)

      ```bash
      npm install -g yo generator-office
      ```

3. In a separate directory, run the following command to create a sample add-in project, which we'll use only to set up the Azure app registration.

      ```bash
      yo office
      ```
      When prompted, choose the following options:
      
      ![image](https://github.com/user-attachments/assets/4a188eff-6b00-490c-9d7b-bde9299cb626)
      
      This will generate an empty sample add-in which we will use to create the app on
      azure. This project will not be used later on.

3. Create the app registration on Azure

      From the folder where the sample project was created, run the following command: `npm run configure-sso` which will prompt you to sign in. 
      Once you sign it, it will create the app registration on Azure automatically. 
      
      Take note of the client ID and the client secret, we will be using them on the
      actual project (copy and store them somewhere).
      
      Once this step is done, your app registration has been created on Azure, you can
      delete the sample project you generated as we won't be needing it anymore.

4. Add required API scopes in Azure

      In your Azure app registration, navigate to API permissions, and click on Add a
      permission -> Microsoft Graph -> Delegated permissions and add the
      permissions listed in the screenshot below
      
      ![image](https://github.com/user-attachments/assets/4ac9e63c-ae44-4d7f-beea-641f584165c1)

### 3. Back to the project directory

1. Create an environment variable file, name it `.ENV`

      .ENV
      ```
      CLIENT_ID=YOUR_CLIENT_ID_HERE
      CLIENT_SECRET=YOUR_CLIENT_SECRET_HERE
      GRAPH_URL_SEGMENT=/me
      NODE_ENV=development
      PORT=3000
      QUERY_PARAM_SEGMENT=
      SCOPE=User.Read 
      ```

2. Replace client ID and secret in the following files

      In `manifest.xml`:
      ```xml
            <WebApplicationInfo>
              <Id>YOUR_CLIENT_ID_HERE</Id>
              <Resource>api://localhost:3000/YOUR_CLIENT_ID_HERE</Resource>
              <Scopes>
                <Scope>User.Read</Scope>
                <Scope>profile</Scope>
                <Scope>openid</Scope>
                <Scope>email</Scope>
                <Scope>Mail.Read</Scope>
                <Scope>Mail.ReadWrite</Scope>
                <Scope>Mail.Send</Scope>
                <Scope>profile</Scope>
              </Scopes>
      ```
      
      In `fallbackauthdialog.ts`
      ```typescript
      const clientId = "YOUR_CLIENT_ID_HERE"; //This is your client ID
      ```




