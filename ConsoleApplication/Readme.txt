This project illustrates how to use CSOM for:
- Reading Project and Tasks
- Adding Tasks, and Timephased Assignment
- Adding Timephased Actual on assignment

It's a simple console application, and work on Project Online, as well as Project On Prem.
All methods are commented, specially where things seems not obvious.

This project is intented to be a helper for your development, because de Project Server 2013 SDK is not fully documented.
On the Support Forum in Technet, a lot of questions are often asked, related to this area of development. So I hope that this project will save (a bit of) your time.

The pre requisite are:
- Having a Project Server 2013 instance on line or On Premise
- Having at list One project
- The periods must be created, and the Timesheet for the current period must be created (simply click on the TimeSheet link on PWA). This issue will be solved soon.
- The current user must be defined as a resource
- On your development machine, VS2013, Project Server 2013 SDK, and the SharePoint Server 2013 Client Components SDK.
Normally, all dependencies are included in the Libraries folder. Some of these assemblies must be in GAC

If you need more info, contact me:
mail: sylvain.gross@neos-sdi.com
twitter: @SylvainGrossNeo

If you have some suggestions, or improvement proposal, don't hesitate to contact me. 
This app may contain some bugs or "undiscovered behaviour" ;-), it's not intended to be used as is in Production. Your feedback are welcome to improve this code.

You can also contact me if you have difficulties to create a Project Online tenant, or Project Server environment. 