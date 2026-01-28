<img src="https://raw.githubusercontent.com/JensKrumsieck/PowerPoint-Remote/master/PowerPoint%20Remote/Resources/PPTRemote.png" alt="PPTRemote" width="250px"/>

# PowerPoint-Remote
Use your mobile device connected to your router (with or without internet) and run your PowerPoint with picture copies of the slideshow on your mobile device. No need to use up system resources by having a live remote picture of your PowerPoint presentation. I use it to lead worship.

Original project: https://github.com/JensKrumsieck/PowerPoint-Remote

Views of the remote on your mobile device:
<img src="https://github.com/AV-Nick317/PowerPoint-Remote-by-AV_Nick/blob/master/.github/PowerPoint%20Remote%20Control.jpeg.png" alt="RemoteWindow" />

Examples of what this can do in YouTube video:
https://youtu.be/b0eDuyEK6JM



<img src="https://github.com/AV-Nick317/PowerPoint-Remote-by-AV_Nick/blob/master/.github/AV_Nick-video1%20-%20frame%20at%200m0s.jpg" alt="RemoteWindow" />
<img src="https://github.com/AV-Nick317/PowerPoint-Remote-by-AV_Nick/blob/master/.github/AV_Nick-video1%20-%20frame%20at%200m15s.jpg" alt="RemoteWindow" />
<img src="https://github.com/AV-Nick317/PowerPoint-Remote-by-AV_Nick/blob/master/.github/AV_Nick-video1%20-%20frame%20at%200m21s.jpg" alt="RemoteWindow" />

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

