# VB Outliner
Brings back Color Coding in the Outlining for VB.NET  
[Also on Visual Studio Gallery](https://visualstudiogallery.msdn.microsoft.com/849848ac-ea59-4815-9892-24e6f7deae57)

![Brings back Color Coding in the Outlining for VB.NET](/VB Outliner/VB Outliner.jpg?raw=true "Brings back Color Coding in the Outlining for VB.NET")

## Disclaimer
I've been missing this feature since VS2008 I think, so I decided to make something.  
It is my first Extension for Visual Studio. It was originally programmed for VS2013, but I adapted it to 2015.  
  
This project is in **BETA** and some elements doesn't work properly, even if I already use it in my workflow.  
**PLEASE CONTINUOUSLY SAVE YOUR WORK IF YOU USE THIS EXTENSION, IT MIGHT UNEXPECTEDLY CRASH**  
  
*`I do not intend to develop this project further, as my knowledge in VS Extensions is really limited.  
Feel free to propose updates to make it better for everyone.`*

## Informations
I didn't know how to change the Default Outlining, so I created a new Outlining on top of the existing one.  
There might be better options.

## Known Issues
* Visual Studio might crash in unexpected cases. I added a lot of Try Catch, but it might still happen.
* There are blank spaces inside the collapsed Outline comments.  
  *The `'''<summary>` tag is removed in the collapsed text, but as the real Outline is still behind, the space needs to be filled to hide it.*
* The highliting of User Types (Classes, Interfaces, Modules, etc. ) is not working in VS2015
* The Tooltip might not come as expected (location or text)
* The Tooltip is displaying white text, even is VS2015 is normally displaying colored Tooltips.  
  *I do not know how to make it colored*

## License

VB Outliner is licensed under the GNU GENERAL PUBLIC LICENSE V3 - the details are in the file [License.md](/LICENSE.md "License.md")  
The goal is to make everyone learn from any possible improvements by sharing it back to the community.  
Thank you.
