# Generate a DLL Base Address

The app randomly generates a memory address which you can specify in your library.


Every ActiveX OCX file and DLL are supposed to have their own DLL Base Address which specifies a location in the memory where it can be loaded. If a program has the same DLL Base Address, windows will move your application to a different location, which obviously damages performance as your app loads up. That is why you should always change it from its default base address by going to Project|Properties|Compile. Once you are there, however, you need to know a valid value to enter! The code below will generate a random base address for you. Read the comments to find out how!


**[Text and code part borrowed from here](http://www.developerfusion.com/code/281/generate-a-dll-base-address/)**