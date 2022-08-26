# cibctoday-o365-personalgreeting

## Summary

The Personal Greeting web part displays the current user's photo, greeting message, display name(first name followed by the last name), and today's date in orders. Adding an optional message under the greeting message is possible if needed. This allows you to configure text message, color, and size for the Greeting message, the Optional message and only color and size for the Date part.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.15-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

> No special pre-requisites needed at this moment.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| cibctoday-o365-personalgreeting-solution | Eunju Han (eunju1.han@cibc.com, CIBC, [Workplace Profile](https://cibc.workplace.com/profile.php?id=100065805794505&sk=about)) |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | Aug 23, 2022     | Update by having a current user's Preferred Name and today's date  |
| 1.0     | Aug 15, 2022     | Initial release - Persona control by having a current user's picture, DisplayName and today's date|

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> If you have spfx-fast-serve and spfx-fast-serve-helpers installed
- in the command-line run:
  - **npm run serve** 

## Deploy
* `gulp clean`
* `gulp build --ship`
* `gulp bundle --ship`
* `gulp package-solution --ship`
* Deploy the `.sppkg` file from `sharepoint\solution` to your tenant App Catalog by using the PS script to have unique permissions for the app. Only an authorized AAD group will have access to manage the app.
* If needed, upload the `.sppkg` file from `sharepoint\solution` to your tenant App Catalog manually for testing purposes
	* E.g.: https://&lt;tenant&gt;.sharepoint.com/sites/AppCatalog/AppCatalog  
* Add the web part to a site collection, and test it on a page

## Features

This extension illustrates the following concepts:

- React
- Office UI Fabric
- spfx-fast-serve

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
- [SPFx Fast Serve Tool](https://github.com/s-KaiNet/spfx-fast-serve) - A command line utility, which modifies your SharePoint Framework solution, so that it runs continuous serve command as fast as possible

