# Offline SP Form
It is a form customizer.

## Summary

This SPFx list form customizer enables users to input offline entries, specifically in areas with no network or intermittent internet access. All submitted items will be stored locally in an IndexedDB. Once connectivity is restored, the locally stored data will be saved to the associated SharePoint list. <br/>
This form customizer is connected to the built‑in Issue Tracker list to demonstrate how you can prevent data loss in areas with weak or unstable network coverage. You can absolutely use it with any other SharePoint list—you just need to define the correct interface for that list and update the list’s GUID in the code.

## Used SharePoint Framework Version


![version](https://img.shields.io/badge/version-1.22.0--rc.1-yellow.svg)

## Solution

| Solution      | Author(s)                                               |
| ------------- | ------------------------------------------------------- |
| offlineSpForm | Shahab Matapour ([LinkedIn profile](https://www.linkedin.com/in/shahabmatapour/))|

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.1     | March 12, 2026   | Update comment  |
| 1.0     | March 10, 2026   | Initial release |

---

## Create a package

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **heft build** 
  - **heft package-solution --production**
- The generated package file (.sppkg) will be located in the `SharePoint/Solutions`

## Features

Description of the extension that expands upon high-level summary above.

This extension illustrates the following concepts:

- topic 1
- topic 2
- topic 3

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
