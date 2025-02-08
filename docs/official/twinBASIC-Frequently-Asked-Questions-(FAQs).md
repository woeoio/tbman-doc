# twinBASIC Frequently Asked Questions

### [General](#general) - [Installation](#installation) - [Using twinBASIC](#using-twinbasic)

## General

### What is twinBASIC?

twinBASIC is a new BASIC language and development environment (IDE) aiming to be 100% backwards compatible with VB6. 

### Who is behind twinBASIC?

twinBASIC is the work of Wayne Phillips, who operates the company [Everything Access](https://www.everythingaccess.com/), a well established provider of professional tools and services for Microsoft Access and VBA generally, including the popular vbWatchdog software.

### Where can I get twinBASIC?

The latest version can be downloaded from the [Releases section](https://github.com/twinbasic/twinbasic/releases) of the [main twinBASIC GitHub repository](https://github.com/twinbasic/twinbasic).

### What is the current status of the project?

twinBASIC is currently late into the **Beta** stage, under development and not yet at a stable 1.0 release. All of the VB6/VBA7 syntax and intrinsic functions have been implemented. All of the basic controls except the OLE control, and about half of the Common Controls, have been implemented. It supports Forms, Classes, and UserControls-- both as compiled OCX/DLL controls and as in-project code (i.e. like .ctl files). However, not all features of these, such as properties, events, and methods, have been completed. Additionally, the `Printer` object, MDI form support, and VBG project group support are not yet implemented. Additionally, there's a fair number of bugs remaining.

However, **tB can already run many existing projects**, even fairly complex and large ones. Many community members have managed to get their apps and other open source apps up and running with little to no difficulty, and created new projects from scratch. Check out these examples for a good real world demonstrate of how far along the project is:\
Krool's [VBCCR](https://github.com/Kr00l/VBCCR) and [VBFlexGrid](https://github.com/Kr00l/VBFLXGRD) controls, Ben Clothier's [TwinBasicSevenZip](https://github.com/bclothier/TwinBasicSevenZip), Carles PV's [Lemmings](https://github.com/fafalone/Lems64), Don Jarrett's [basicNES](https://github.com/fafalone/basicNES) Nintendo emulator, and Jon Johnson's [ucShellBrowse/ucShellTree](https://github.com/fafalone/ShellControls), [FileActivityMon ETW Event Tracer](https://github.com/fafalone/EventTrace), [cTaskDialog](https://github.com/fafalone/cTaskDialog64), and [many more](https://github.com/fafalone).

### Is there an estimated timeline for when expected features will become available?

Yes, see the [twinBASIC Roadmap](https://github.com/twinbasic/twinbasic/issues/335) in the Issues section for the latest update to the timeline. This roadmap only covers major components; smaller features are implemented in a less formal manner, usually when the related part of the codebase is being worked on.

### What new features does twinBASIC have compared to VB6?

**Many!** It has 64bit compilation (using VBA7x64 compatible syntax), generics, overloading, multithreading (API-only right now, built in syntax coming soon), inheritance, ability to define interfaces and coclasses in your project using BASIC-style syntax, Unicode support in all controls and the editor (.twin files only), support for modern image formats, numerous enhancements to *Implements*, ability to create standard DLLs and kernel mode drivers, ability to set UDT packing alignment, and dozens of others, all available *right now*, with many more planned in the future.

For a full list of all the new features available right now, see the Wiki article [Overview of features new to twinBASIC](https://github.com/twinbasic/documentation/wiki/twinBASIC-Features).

### Where can I learn more about twinBASIC, find documentation, and participate in the community?

[twinBASIC Home Page](https://twinbasic.com)

twinBASIC GitHub: [Main section](https://github.com/twinbasic/twinbasic) | [Issues](https://github.com/twinbasic/twinbasic/issues) | [Discussions](https://github.com/twinbasic/twinbasic/discussions) | [Language Design](https://github.com/twinbasic/lang-design) | [ Language Specification](https://github.com/twinbasic/lang-spec) | [Documentation](https://github.com/twinbasic/documentation/wiki)

[twinBASIC Discord](https://discord.gg/UaW9GgKKuE)

[twinBASIC Forum on VBForums](https://www.vbforums.com/forumdisplay.php?108-TwinBASIC)

### Is twinBASIC Open Source?

While open source models are possible in the future, at this time the compiler is not. There are plans in the works to open source the IDE. To address some of the major concerns this presents, once tB hits it first major release, the source code will be placed in escrow, to be released to the community in the event the author disappears or is unable to continue working on it due to death or serious illness/injury. 

### How much does twinBASIC cost? 

There are 3 editions of twinBASIC: The Community Edition is FREE. A splash screen is placed on compiled 64bit binaries and certain features like advanced optimized compilation and future cross-platform compilation are unavailable, but there are no restrictions on core language features or royalties imposed. To get those features, subscriptions are available for the Professional and Ultimate editions. For more details, including current pricing for Professional and Ultimate editions, [see this page](https://twinbasic.com/preorder.html).

Note that you can change the subscription level at any time and the community edition is always available. There will be no lockout (see [the previous statement regarding escrow](https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs)#is-twinbasic-open-source)) so you will always have the ability to develop, test and compile. 

### Can I pay a one-time fee for a perpetual license?

Due to the need for continuing income to be able to develop twinBASIC, subscriptions are the primary model for premium versions, which [are available](https://twinbasic.com/preorder.html) on a month-to-month or yearly basis. However, right now for a limited time, a buy-once perpetual license is available in the form of the [VIP Gold Lifetime Licence Initiative](https://twinbasic.com/vip.html). This provides not only a lifetime license to twinBASIC including updates and new versions, but numerous additional benefits available only to people who purchase this license. 


### Can twinBASIC be used to develop commercial products, and what royalties are owed?

There are no restrictions on any edition of twinBASIC; they can all be used to develop commercial products, on a ROYALTY FREE basis. Nothing is owed for selling programs or other products created with twinBASIC. The twinBASIC software itself, however, may not be redistributed without appropriate license.

### What does '100% backwards compatible' mean, technically?

Backwards compatibility refers to matching all publicly documented syntax, included controls, component and control behavior, and control appearance. It does not include undocumented, propietary internal implementation details. So for example all language keywords, functions, and methods will are present and should give the same results, and Forms/Classes/UserControls should implement all the same publicly documented interfaces, but twinBASIC exe files are not internally structured in the same way, and there isn't compatibility with the undocumented VB project info structures in the exe, whose contents have been reverse engineered over the years by the community. 

Currently, all basic controls have reimplementations in twinBASIC that support Unicode and 64bit compilation except the OLE control; and a number of the primary Common Controls have also been reimplemented. Eventually, all controls shipped with VB6 Enterprise Edition will be reimplemented. Until that time, the original controls will still work in 32bit builds, and community members provide some alternatives, for example Krool's VBCCR controls and VBFlexGrid control all work and have 64bit-compatible twinBASIC versions.

### So some of my projects won't work?

Most projects do not use these reverse engineered internals, but some do: mostly commonly, for self-subclassing and callbacks inside Forms/Classes/UserControls; and also for multithreading and inline assembly. These routines have native support in twinBASIC without requiring internals hacks, so replacing these small parts of a few programs is very simple: `AddressOf` is supported on class members, so you can use regular subclassing and callback methods as you would if they were in a .bas module. `CreateThread` can be called without any special steps. And tB supports statically linked .obj files allowing incorporation of code from other languages, and further support is expected in the future.

Additionally, twinBASIC redirects the most common msvbvm60.dll (also msvbvm50.dll/vbe6.dll/vbe7.dll) functions used as `Declare` statements by users to internal implementations, all of which also work in x64, if you add the `PtrSafe` keyword like any other DLL definition. The following functions currently have redirects: `VarPtr, GetMem1, GetMem2, GetMem4, GetMem8, PutMem1, PutMem2, PutMem4, PutMem8, __vbaObjSet, __vbaObjSetAddRef, __vbaObjAddRef, __vbaCopyBytes, __vbaCopyBytesZero, __vbaRefVarAry`, and `__vbaAryMove`. You may continue to use these with `Declare` statements to support the particular signatures you prefer. Further, declares for olepro32.dll are redirected to identical functions in oleaut32.dll, as olepro32 was deprecated by NT4 and doesn't have a 64bit version.


Other than those special cases, it's exceedingly rare for projects to depend on reverse engineered internals. So the vast majority of projects run with zero modification. 

### How do I report bugs or other problems?

The best way is to [create an issue](https://github.com/twinbasic/twinbasic/issues) in the twinBASIC GitHub repository.

You can also create a post in the #bugs channel of the [twinBASIC Discord server](https://discord.gg/UaW9GgKKuE).

### Is the twinBASIC IDE available in other languages?

There's currently no official translations for other languages, but some additional languages are expected soon as many members have expressed interest in this and generously offered to provide translations. 

## Installation

### What are the requirements for twinBASIC?

twinBASIC is supported on Windows 7 through Windows 11. The installation is portable; you need only to extract the downloaded zip file then run; there's no installer to run. 

WebView2 is required. This is normally preinstalled on newer versions of Windows, and is installed along with Edge if you've installed that browser. You can also obtain it from [Microsoft's website](https://developer.microsoft.com/en-us/microsoft-edge/webview2?form=MA13LH#download-section). Select the Standalone Evergreen x86 version:

![image](https://github.com/twinbasic/documentation/assets/7834493/94490c87-fafe-4d5b-ae39-d3cedba1c21d)


### twinBASIC won't run; says there's an invalid entry point.

This issue is sometimes encountered on Windows 7. To be used on Windows 7, the OS must be fully updated; this error results from one or more missing updates. Run Windows Update to make sure you have all recent updates installed. If you still have problems, you can drop by the Discord or submit an issue on GitHub (see [`How do I report bugs or other problems?`](https://github.com/twinbasic/documentation/wiki/twinBASIC-Frequently-Asked-Questions-(FAQs)#how-do-i-report-bugs-or-other-problems))

### How do I install twinBASIC?

tB does not require a full installation process, you need only extract the ZIP file. Download the latest version from the [Releases page](https://github.com/twinbasic/twinbasic/releases), named `twinBASIC_IDE_BETA_xxx.zip` (where xxx is a version number; click on 'Assets' to expand the file list if it's not already visible), and extract the contents of the zip to a folder of your choice. It will run from this folder, not copy files elsewhere.  

> [!IMPORTANT]
> It's highly recommended that you either install each new version to a new folder, or first delete all existing files in the destination folder. Unusual bugs have been traced to simply attempting to overwrite old installations with new ones.


### How big is the twinBASIC installation?

The IDE is quite small, it's currently only 22MB, and that is half due to LLVM libraries.

### Where is twinBASIC IDE data stored?

In addition to the directory you extract the IDE to, twinBASIC creates a folder in `%APPDATA%\Local\twinBASIC` and stores some settings in the Registry under `HKCU\Software\VB and VBA Program Settings\twinBASIC_IDE`.

### Is twinBASIC safe? (Some scanner) says it's malicious.

Anyone who has ever tested their own programs against a wide variety of AV engines knows that unless your exe is 64bit and signed with a high-level certificate (and maybe not even then, until it's manually added to a trust list), false positives in a small number are simply a way of life. twinBASIC's IDE and compiler executables, like all apps in its position, may trigger a small number of positives on services like VirusTotal, particularly 32bit apps. These are almost always not from major vendors and/or "AI" based algorithmic detection.

## Using twinBASIC

### How do I import my VB6 project into twinBASIC?

The easiest way is through the import wizard. When you first start the twinBASIC IDE, you're presented with the New Project dialog- this contains an 'Import from VBP' option:

![image](https://github.com/twinbasic/documentation/assets/7834493/7e1cb69c-6db3-4f3f-aea1-c1fae25938a2)

This is currently the only way to import Forms, UserControls, and Resources. For standard modules and class modules, you can import them individually by right-click the 'Sources' folder in the Project Explorer, and choosing Import file...:

![image](https://github.com/twinbasic/documentation/assets/7834493/60335cfc-3573-489a-90e9-9dbec2b2113c)

>[!NOTE]
>If you import Forms or UserControls through this method, they are currently processed as plain text and will not be recognized by the compiler. This will be fixed in the future. For now, please import these as part of a VBP project.

### Does twinBASIC support addins?

Addins for VB6 and VBA are not supported by the twinBASIC IDE. However,  tB has it's own powerful addin infrastructure based on modern web technologies. See Samples 10 through 16 in the 'Samples' tab of the New Project dialog. 

twinBASIC supports **creating** addins for VBA. It's currently the only tool that supports creating these addins for 64bit Office using a language with 100% compatible syntax. See Sample 4 and Sample 5. 

### How do I use resources in twinBASIC?

Currently tB does not have a dedicated source editor; instead, resources are managed through the Project Explorer. In the tree, you'll see a Resources folder; by default, it will include ICON in a Standard EXE, and MANIFEST, if you've chosen to enable Visual Styles:

![image](https://github.com/twinbasic/documentation/assets/7834493/71ddde83-a091-47e3-b5b8-681954b0639d)

You can create additional folders here, using their standard names. For example a BITMAP group could be added, then used with `LoadResImage`. Unlike its predecessor, tB does not restrict the type of resources: you can create any type of folder you want, and import binary data into it. For example, some community projects have inserted `UIFILE` resources for Ribbon controls and `DIALOG` resources for property sheets. Resources can be imported by right-clicking the folder you want them in, and selecting Add->Import file... from the menu.

If you're importing a project, the resources in a linked .res file will be imported automatically.

#### Strings

String table resources are currently treated specially; they're edited in the IDE as JSON. If you import from VBP with a .res, string resources will be automatically converted. If you right click the 'Resources' folder, and go to the 'Add' submenu, at the bottom, you'll find "Add resource: String table" that adds one populated with example strings:

![image](https://github.com/twinbasic/documentation/assets/7834493/97cc8655-7a8b-47f3-b52c-eb1ddfce662f)

#### Group names

If you create a new folder for a standard resource type, twinBASIC currently recognizes the following names, which you should use to create a folder under Resources:

BITMAP\
CUSTOM\
CURSOR\
ICON\
MANIFEST\
RCDATA\
STRING

For other standard types, you must use the # (pound sign) followed their number. For example, for DIALOG (RT_DIALOG) resources, do not name the folder dialog, it must be named `#5`. ANICURSOR would be named `#21`. And so on, for the [standard types](https://learn.microsoft.com/en-us/windows/win32/menurc/resource-types) with `RT_` constants. For any others, you can use any name you want, e.g. UIFILE can just be named UIFILE. 


>[!NOTE]
>At this time, .res files can only be imported as part of a VBP.

### How do I set my own icon for my program?

By default, imported projects use whichever icon (or no icon) was used previously, and newly created projects use the twinBASIC logo. For either, you can of course set a new icon. If you're not already familiar with using resources in twinBASIC, see the FAQ entry right above this one. The icon used for your application in Explorer is the one in the Resources\ICON folder that comes first alphabetically. If you do not have an ICON folder, you can create one by right-click and selecting Add->Add folder.

![image](https://github.com/twinbasic/documentation/assets/7834493/8611d12a-d7a6-48cc-9544-cb27c5299aa5)

In the above picture, MyOwnIcon.ico would be used by Explorer and other apps to represent your .exe, as it comes before twinBASIC.ico alphabetically. 

Note that this is not the icon of your form; icons for forms are set by the "Icon" property in the Properties list.