# outlook-exchange-oauth-connector

This program connects to an Exchange service (e.g., Outlook Office 365) using OAuth2 authentication and impersonates a user to access their email data.

## Prerequisites

Before running the program, make sure you have the following:

- **Microsoft Outlook App** registered in Azure AD (for OAuth authentication)
- **App ID** (Client ID), **Client Secret**, and **Tenant ID** for the registered application
- **.NET Core** or **.NET Framework** environment set up

## Required NuGet Packages

Ensure that the following NuGet packages are installed in your project:

- `Microsoft.Identity.Client` - for handling OAuth authentication.
- `Microsoft.Exchange.WebServices` - for connecting to the Exchange service.

You can install them using the following commands:

```bash
dotnet add package Microsoft.Identity.Client
dotnet add package Microsoft.Exchange.WebServices
```

## Configuration
#### Set up your Azure AD Application:

Register your app in Azure AD and get the following values:
- OUTLOOK_APPID (Client ID)
- OUTLOOK_SECRETID (Client Secret)
- OUTLOOK_TENANTID (Tenant ID)
- OUTLOOK_SCOPES (Required scopes, e.g., https://outlook.office365.com/.default)

### Update the code with your credentials:

  Replace OUTLOOK_APPID, OUTLOOK_SECRETID, OUTLOOK_TENANTID, and OUTLOOK_SCOPES with your actual values in the code.

### Running the Program
1. Clone the repository or download the source code.
2. Open the project in your preferred IDE (e.g., Visual Studio, Visual Studio Code).
3. Replace the placeholders (OUTLOOK_APPID, OUTLOOK_SECRETID, OUTLOOK_TENANTID, OUTLOOK_SCOPES) with your actual Outlook application credentials.
4. Call the ConnectToExchangeService method, passing the email address you want to impersonate, to connect to the Exchange service.

### Example

``` C# code
string emailAddress = "user@example.com";
bool isConnected = await ConnectToExchangeService(emailAddress);
if (isConnected)
{
    Console.WriteLine("Successfully connected to Exchange service.");
}
else
{
    Console.WriteLine("Failed to connect to Exchange service.");
}

```

## Error Handling
- Any errors during token acquisition or Exchange service connection are logged to the console.
- The ConnectToExchangeService method will throw an exception if an error occurs.