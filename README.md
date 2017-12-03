# VB Outliner
[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=W3JMUBMSWUKNY)

Brings back Color Coding in the Outlining for VB.NET
[Also on Visual Studio Gallery](https://visualstudiogallery.msdn.microsoft.com/849848ac-ea59-4815-9892-24e6f7deae57)

![Brings back Color Coding in the Outlining for VB.NET](VB%20Outliner/VB%20Outliner.jpg "Brings back Color Coding in the Outlining for VB.NET")

## Disclaimer
I've been missing this feature since VS2008 I think, so I decided to make something.  
It is my first Extension for Visual Studio.

I do not have VS2015 to test it but works perfectly on VS2017.
  
*`I do not intend to develop this project further, as my knowledge in VS Extensions is really limited.  
Feel free to propose updates to make it better for everyone.`*

## Informations
I didn't know how to change the Default Outlining, so I created a new Outlining on top of the existing one.  
There might be better options.

## Known Issues
* There are blank spaces inside the collapsed Outline comments.  
  *The `'''<summary>` tag is removed in the collapsed text, but as the real Outline is still behind, the space needs to be filled to hide it.*

## License

VB Outliner is licensed under the GNU GENERAL PUBLIC LICENSE V3 - the details are in the file [License.md](/LICENSE.md "License.md")  
The goal is to make everyone learn from any possible improvements by sharing it back to the community.  
Thank you.
