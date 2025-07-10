<img src="https://raw.githubusercontent.com/JensKrumsieck/PowerPoint-Remote/master/PowerPoint%20Remote/Resources/PPTRemote.png" alt="PPTRemote" width="250px"/>

# PowerPoint-Remote (a remix of a previous app)

Orignianal project: https://github.com/JensKrumsieck/PowerPoint-Remote

[https://github.com/FirstPet31415/PowerPoint-Remote-by-Sylvan-Finger/blob/master/.github/PowerPoint%20Remote%20Control.jpeg.png ](https://github.com/AV-Nick317/PowerPoint-Remote-by-AV_Nick/blob/master/.github/PowerPoint%20Remote%20Control.jpeg.png)

This is a program to give you a picture of what is similar to what PowerPoint presentation view on you computer, but presents a copy of it on your mobile device. 

Usernotes: I have not found the original QR code to always work. Please find your PC's IP address and add the port ":3000" afterwards, as shown in the image above so that you can get the mobile device browser connected to the right address. 

Manually export your PowerPoint presentation using the included macro for PowerPoint with the path link updated to wherever you have installed your PowerPoint Remote app. 

Please contact me for any help needed for getting started. 

This is supposed to run like the program Clicker, but doesn't use up your PC's system resources while running. 



Notes from previous project:

**A simple WPF/Web based Remote for Powerpoint.**

<img src="https://raw.githubusercontent.com/JensKrumsieck/PowerPoint-Remote/master/.github/screenshot.png" alt="PPTRemote" width="250px"/>

Open a PPT File to start the Server and use the QR Code to navigate your webbrowser to the internal server. You will see a preview of the current slide and next/previous buttons to navigate through the presentation.

The Slideshow is controlled by sending webrequest from the web interface to the internal server. The server forwards the commands via COM to the powerpoint instance. Screenshots of the Slideshow-Monitor are send as MemoryStream-based response to the /preview request to show the current slide as image.

<sub>some inspiration came from: https://github.com/PuZhiweizuishuai/PPT-Remote-control</sub>

