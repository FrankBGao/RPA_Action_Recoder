# RPA_Action_Recoder
This project is a part of the RWTH-Aachen PADS group's RPA research. This project is for collecting the user-computer's interactions, such as mouse clicks, keyboard keydown, window lanuch and close.

This project is done by C#. This code is calling the win32 for collecting these events.  example_result.json is an example for the collected data.

```javascript
	{
		"Position": "883|454",
		"Event Type": "Mouse",
		"Timestamp": "2018-09-26 15:13:44.700",
		"Event Info": "WM_LBUTTONDOWN",
		"Job id": "b35f0a38-6f2e-4c6c-b528-4887ed778099",
		"Resource": "Human"
	},
	{
		"Job id": "b35f0a38-6f2e-4c6c-b528-4887ed778099",
		"Resource": "Human",
		"Event Info": "LeftCtrl",
		"Timestamp": "2018-09-26 15:13:46.536",
		"Event Type": "Keyboard"
	},
	{
		"Event Type": "Keyboard",
		"Timestamp": "2018-09-26 15:13:46.719",
		"Event Info": "C",
		"Job id": "b35f0a38-6f2e-4c6c-b528-4887ed778099",
		"Resource": "Human"
	},
	{
		"Job id": "b35f0a38-6f2e-4c6c-b528-4887ed778099",
		"Resource": "Human",
		"Event Info": "https://svn.win.tue.nl/trac/prom/wiki/GettingStarted ",
		"Timestamp": "2018-09-26 15:13:46.719",
		"Event Type": "Clipboard"
	}
```

This is a screenshot of this software.

![alt text](https://raw.githubusercontent.com/FrankBGao/RPA_Action_Recoder/master/screenshot.JPG)


Firstly, press the button New Job, it will generate a new job id.

Secondly, press the button Start, it will record the events.

Thirdly, press the button Stop, it will stop recording.

In the last, press the button Into Log File, it will dump the events into a log file.


 
