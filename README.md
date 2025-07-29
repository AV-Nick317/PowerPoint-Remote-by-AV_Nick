<img src="https://raw.githubusercontent.com/JensKrumsieck/PowerPoint-Remote/master/PowerPoint%20Remote/Resources/PPTRemote.png" alt="PPTRemote" width="250px"/>

# PowerPoint-Remote (a remix of a previous app)

Original project: https://github.com/JensKrumsieck/PowerPoint-Remote

https://github.com/AV-Nick317/PowerPoint-Remote-by-AV_Nick/blob/master/.github/PowerPoint%20Remote%20Control.jpeg.png 

This is a program to give you a picture copy of your PowerPoint presentation on your remote device. 

Manually export your PowerPoint presentation using the included macro for PowerPoint with the path link updated to wherever you have installed your PowerPoint Remote app. 

Please contact me for any help needed for getting started. 

This is supposed to run like the program Clicker, but doesn't use up much of any of your devices' system resources while running. 



Notes from previous project:

**A simple WPF/Web based Remote for Powerpoint.**

<img src="https://raw.githubusercontent.com/JensKrumsieck/PowerPoint-Remote/master/.github/screenshot.png" alt="PPTRemote" width="250px"/>

Open a PPT File to start the Server and use the QR Code to navigate your web browser to the internal server. You will see a preview of the current slide and next/previous buttons to navigate through the presentation.

The Slideshow is controlled by sending webrequest from the web interface to the internal server. The server forwards the commands via COM to the PowerPoint instance. Screenshots of the Slideshow-Monitor are send as MemoryStream-based response to the /preview request to show the current slide as an image.

<sub>some inspiration came from: https://github.com/PuZhiweizuishuai/PPT-Remote-control</sub>

