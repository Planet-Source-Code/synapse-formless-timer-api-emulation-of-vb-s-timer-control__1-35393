
Formless Timer: API emulation of VB's Timer control.
	Written by Jahufar Sadique
---------------------------------------------------

	This project contains an ActiveX DLL that will emulate the functionality of VB's Timer control. This sorta thing could come in handy if you ever need to have a timer in a class module (if you are are building an ActiveX DLL). The project also demonstrates a use of an 'EventSink' to get feedback from a function in a BAS module.

	TIMER.DLL should be in the \control folder. If you do not find it (PSC might have removed it), please open TIMERLIB.VBP and compile the DLL. Make sure you have Version Compatibility set to 'Binary Compatibility'. There is demo project in the \demo folder. The \doc folder contains a DOC file that documents the functionality of the DLL.

	You can of course use the 2 class files and the module found in the DLL directly into your own project thereby reducing a runtime dependancy - but doing so will make your code a bitch to debug because of the AddressOf callback thats used.

	I would recommend that you go using the DLL until you have all the glitches worked out of your code and then directly inject the DLL internals into your final build. The DLL is pretty small: 24K, so I don't think it's that big of an runtime overhead. 

	But if you really want to eliminate this, read the below carefully:

	If you mess around with class files by adding them to Standard EXE, VB will automatically alter the class's Instancing property to 'Private' (1). If you add them back into the DLL, please make sure you have the Instancing properties of:

		clsTimer -> set to 'MultiUse' [5]
		clsTimerEventSink -> set to 'Private' [1]

	Also, if you decide to add the class files to your project directly, you will have to change the scope of the following method and property to be 'Public' (they are 'Private' in the DLL):

		TimerID [Property Get]
		StopTimer() [Function]

	Make *sure* you call StopTimer() before your program exits. In the DLL this will be called automatically via the _Terminate 
event - this is why StopTimer() and TimerID was declared as 'Private' in the DLL. Never use the IDE's stop button or the 'End' keyword (using 'End' is bad programming practice anyway). If you do, the IDE will crash.

	Thats about it I think :) Have fun! Hope you find my code useful and easy to understand (I hope my commenting was adequate). 
Vote if you want - but do leave your feedback.

-Jahufar Sadique [jahufar@fastmail.fm]
Colombo, Sri Lanka
29/05/2002

