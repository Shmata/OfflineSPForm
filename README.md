# Offline SP Form
It is a form customizer.

## Summary
In many fields, especially construction, inspectors often need to visit job sites and record information directly into a SharePoint list. Because network coverage on construction sites can be unreliable, there’s a real risk of losing data while filling out forms. This form customizer is designed to prevent that problem by allowing inspectors to keep working even when the connection drops, and then automatically syncing their entries once the network is restored.
<br/>
This SPFx list form customizer enables users to input offline entries, specifically in areas with no network or intermittent internet access. All submitted items will be stored locally in an IndexedDB. I used the [Dexie library](https://www.npmjs.com/package/dexie) as a lightweight wrapper around IndexedDB in this project. Once connectivity is restored, the locally stored data will be saved to the associated SharePoint list. <br/>
This form customizer is connected to the built‑in 'Issue Tracker' list to demonstrate how you can prevent data loss in areas with weak or unstable network coverage. You can absolutely use it with any other SharePoint list—you just need to define the correct interface for that list and update the list’s GUID in the code.

![Network indicator](https://github.com/Shmata/OfflineSPForm/blob/main/src/extensions/offlineSpForm/assets/2.png)

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
Creating and installing this package works the same way as any other SPFx Form Customizer. There’s nothing unique or special about its packaging process, and the standard SPFx build, bundle, and deploy steps apply. For additional guidance, the [Microsoft documentation on SPFx extensions](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-form-customizer#deployment-of-your-extension) covers the full workflow clearly and is the best reference to follow. <br/> 

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **heft build** 
  - **heft package-solution --production**
- The generated package file (.sppkg) will be located in the `SharePoint/Solutions`

## Operating Procedure

After you deploy the form customizer and link it to your SharePoint list, just open the list and select New item—your custom form will load automatically.
![New form](https://github.com/Shmata/OfflineSPForm/blob/main/src/extensions/offlineSpForm/assets/1.png)
<br/><br/>
The default SharePoint list form appears along with a built‑in network indicator.
<br/><br/>
![No Network](https://github.com/Shmata/OfflineSPForm/blob/main/src/extensions/offlineSpForm/assets/3.png)
<br/><br/>
If your connection becomes weak or drops entirely, the indicator reflects that change. <br/><br/>
![Fillout form](https://github.com/Shmata/OfflineSPForm/blob/main/src/extensions/offlineSpForm/assets/4.png)
<br/><br/>
You can continue filling out the form normally—just leave it as is when you're done. Once your network connection returns, the form automatically syncs and saves your item back to SharePoint.
<br/><br/>
![Synce](https://github.com/Shmata/OfflineSPForm/blob/main/src/extensions/offlineSpForm/assets/5.png)
<br/><br/>
It works on new and update forms. 

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Deployment of your Extension](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/building-form-customizer#deployment-of-your-extension)

## ⭐ Final Recommendation

If you found this project useful, give it a star ⭐ and become **0.1% cooler instantly**. <br />
It's scientifically unproven… but widely believed.
