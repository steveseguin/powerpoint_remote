# PowerPoint Remote

With this code and guide, you can remotely control PowerPoint. It's primarily designed for VDO.Ninja remote guests to control a presented PowerPoint slide, but it can be used in other ways as well.

We will be using Autohotkey, a virtual MIDI driver, and optionally VDO.Ninja to control Power Point.  I saw optionally use VDO.Ninja, as you can replace it with a locally connected StreamDeck or other local MIDI-compatible software.

I've only tested this with Windows 11 currently; Apple/Linux users will likely have to tweak the guide as needed.

## Basic setup for the host system

The host system is what is running the desktop version of Power Point.

The host system also needs a virtual MIDI loopback device needs to be installed. I use LoopMIDI for Windows, but there are other options. https://www.tobias-erichsen.de/software/loopmidi.html

The next critical bit of setup for the host system is having Autohotkey v2 installed; https://www.autohotkey.com/download/

The script we want AutoHotKey to run is contained in this code repository, so we need to download and run that code as well;

Download and extract: https://github.com/steveseguin/powerpoint_remote/archive/refs/heads/main.zip

Once extracted, we can just run the `run.ahk` file. Double clicking on it normally does it.

![image](https://user-images.githubusercontent.com/2575698/210496360-074338fe-13e2-4c97-8932-ae58286826ac.png)

With it running, we should see a little green H icon in our task tray.  If we right-click on it, we can select MIDI Input Devices and then select our loopMIDI device.  If we did that correctly, there should be a little checkmark next to it in the menu, indicating our AutoHotKey script is listening to that MIDI device for commands.

![image](https://user-images.githubusercontent.com/2575698/210497319-0a8a45e0-6c89-4ae6-9c1d-e8973500c4c9.png)

At this point, there are a few different ways to use this setup to control PowerPoint; we'll cover some scenarios and their setup below.

## VDO.Ninja guest remote control

A VDO.Ninja remote guest, viewer, or even a publisher can use a built-in control interface to issue next or previous slide commands to the host system.

#### Host setup in VDO.Ninja

This link is used by the host PowerPoint system. 
`https://vdo.ninja/?director=myRoom123&midiin`

`&midiin` tells the client that its forwarding inbound commands to all local MIDI devices.  We can specify specific MIDI devices if multiple are installed, but generally most people will only have the virtual loopMIDI device installed, so its not an issue.

`?director=myRoom123` can be replaced with other options, but in this case we assume the host is the stream director.  They will likely be screen sharing the local Power Point presentation to everyone in the room, and a room director has more control options than a normal user would.

#### Remote client setup in VDO.Ninja

Since we want to use the built-in remote controller for the guest, we will need to be using VDO.Ninja v22.12 or newer. 
`https://vdo.ninja/alpha/?room=myRoom123&ppt`

`&ppt` is what enables the Power Point controls for this guest, which can be see as a left and right arrow.

![image](https://user-images.githubusercontent.com/2575698/210496086-d9e6a1bf-145b-4859-8765-0f5da8ccbae5.png)

`&room=myRoom123` just has the guest join the specific room.

## Remote StreamDeck control via MIDI

StreamDecks can send MIDI commands

#### Host setup in VDO.Ninja

This link is used by the host PowerPoint system. 
`https://vdo.ninja/?director=myRoom123&midiin`

`&midiin` tells the client that its forwarding inbound commands to all local MIDI devices.  We can specify specific MIDI devices if multiple are installed, but generally most people will only have the virtual loopMIDI device installed, so its not an issue.

`?director=myRoom123` can be replaced with other options, but in this case we assume the host is the stream director.  They will likely be screen sharing the local Power Point presentation to everyone in the room, and a room director has more control options than a normal user would.

#### Remote client setup in VDO.Ninja

We can install loopMIDI on the remote StreamDeck machine and install a MIDI plugin onto our StreamDeck.  From there, we can output the MIDI control-change command 110 with value 10 and 11 from our StreamDeck to the loopMIDI virtual loopback device.  VDO.Ninja will detect it when using `&midiout`, and forward it to the host system. 

`https://vdo.ninja/?room=myRoom123&midiout`

`&room=myRoom123` just has the guest join the specific room.

## Local StreamDeck control via MIDI

You don't need VDO.Ninja in this case, just StreamDeck and the MIDI plugin. Outputting the MIDI control-change command 110 with value 10 and 11 from our StreamDeck 
to the virtual MIDI loopback device is all we need to do.

In our AutoHotKey app, assuming we have the MIDI loopback device selected as the input device, the StreamDeck will be sending commands to the script directly via MIDI. This of course only works locally; not remotely.

## Local Web control via MIDI

You don't need to technically use a StreamDeck to issue MIDI commands; browser have something called WebMIDI, which is what VDO.Ninja uses. Making your own little web app in the browser that issues MIDI commands to the local MIDI device will give web developers quite a bit of flexibility. It's out-of-scope of this guide, but it's something that's easy enough to look up how to do.

## Remote Web control via VDO.Ninja

VDO.Ninja has an IFRAME API, which will let you issue commands to VDO.Ninja via a parent window. In this way, you can leverage both the remote peer to peer power of VDO.Ninja, but also the MIDI functionally. A sample web app is provided, so if you want to customize the VDO.Ninja controller or embed the VDO.Ninja controller into your app, it's fairly easy to do.

## HTTP / Websocket API

Lastly, VDO.Ninja has an HTTP / WSS API, and it lets you issue commands to the host system.  HTTP commands work with StreamDecks, so you can have a remote client issue commands to your VDO.Ninja host system, controlling your PowerPoint, without even that client having VDO.Ninja open themselves.  The reason we'd be using VDO.Ninja is just to make use of the MIDI code that's already there.  

We'd just need to have `https://vdo.ninja/?api=YOURAPIKEY&midiin` open to make use of this as a host- we don't need to be even viewing or sharing a video.  The API key can be made up, so long as it matches the HTTP request.

`https://api.vdo.ninja/YOURAPIKEY/nextSlide` and `https://api.vdo.ninja/YOURAPIKEY/prevSlide` woudl be the HTTP commands. Refer to the API documentation for details on the WSS commands.
