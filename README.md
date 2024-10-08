# PowerPoint Remote (works with more than just powerpoint)

With this code and guide, you can remotely control PowerPoint, OBS, or anything really, without needing websockets, SSL, or having to open insecure ports. It's primarily designed to let VDO.Ninja remote guests remote control a presented PowerPoint slide, but it can be used in other ways as well.

![image](https://user-images.githubusercontent.com/2575698/210701455-8900148e-3f95-496f-9565-5226c95884f3.png)

Video guide: [https://youtu.be/ORH8betTt8Y](https://youtu.be/ORH8betTt8Y)


In our main example will be using Autohotkey v2, a virtual MIDI driver, and optionally VDO.Ninja to control Power Point.  I say 'optionally VDO.Ninja', as you can replace it with a locally connected StreamDeck or other local MIDI-compatible software.

I've only tested this with Windows 11 currently; Apple/Linux users will likely have to tweak the guide as needed.

The AutoHotKey script can be modified to work with other apps, not just PowerPoint. I'll add more apps as they are requested.  You can also use alternatives to AutoHotkey, including the OBS plugin [obs-midi-mg](https://github.com/nhielost/obs-midi-mg), which will let you control OBS with MIDI directly, without the need for AHK. This is a great way to control OBS remotely online without the need of opening an secure port on your computer or having to enable SSL to use from the web.

## Basic setup for the host system

The host system is what is running the desktop version of Power Point.

The host system also needs a virtual MIDI loopback device needs to be installed. I use LoopMIDI for Windows, but there are other options. https://www.tobias-erichsen.de/software/loopmidi.html

The next **critical** bit of setup is for the host system to have **have Autohotkey v2 installed**; https://www.autohotkey.com/download/

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

`&midiin` tells the client that its forwarding inbound commands to all local MIDI devices.  We can specify specific MIDI devices if multiple are installed, but generally most people will only have the virtual loopMIDI device installed, so its not an issue. For more details, refer to: https://docs.vdo.ninja/advanced-settings/api-and-midi-parameters/midi

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

We can install loopMIDI on the remote StreamDeck machine and install a MIDI plugin onto our StreamDeck (ie: https://apps.elgato.com/plugins/se.trevligaspel.midi).  From there, we can output the MIDI control-change command 110 with value 10 and 11 from our StreamDeck to the loopMIDI virtual loopback device.  VDO.Ninja will detect it when using `&midiout`, and forward it to the host system. 

`https://vdo.ninja/?room=myRoom123&midiout`

`&room=myRoom123` just has the guest join the specific room.

Refer to the VDO.Ninja documentation on MIDI Pass-Thru mode for more detail: https://docs.vdo.ninja/advanced-settings/api-and-midi-parameters/midi#midi-pass-through-mode

## Local StreamDeck control via MIDI

You don't need VDO.Ninja in this case, just StreamDeck and the MIDI plugin. Outputting the MIDI control-change command 110 with value 10 and 11 from our StreamDeck 
to the virtual MIDI loopback device is all we need to do.

In our AutoHotKey app, assuming we have the MIDI loopback device selected as the input device, the StreamDeck will be sending commands to the script directly via MIDI. This of course only works locally; not remotely.

## Local Web control via MIDI

You don't need to technically use a StreamDeck to issue MIDI commands; browser have something called WebMIDI, which is what VDO.Ninja uses. Making your own little web app in the browser that issues MIDI commands to the local MIDI device will give web developers quite a bit of flexibility. It's out-of-scope of this guide, but it's something that's easy enough to look up how to do.

## Control OBS with MIDI, bypassing the need for AHK

The [obs-midi-mg](https://github.com/nhielost/obs-midi-mg) project is a customizable plugin for OBS that will let you control OBS via MIDI.  This means it should work with VDO.Ninja's MIDI controls, bypassing the need for websockets, firewalls, SSL, and other security considerations when remotely controlling OBS from the web.

https://github.com/nhielost/obs-midi-mg/releases/tag/v2.3.0 is a version of the plugin that was last confirmed to be working well.

## Remote Web control via VDO.Ninja (IFRAME API)

VDO.Ninja has an IFRAME API, which will let you issue commands to VDO.Ninja via a parent window. In this way, you can leverage both the remote peer to peer power of VDO.Ninja, but also the MIDI functionally. A sample web app using the IFRAME API can be found here, https://vdo.ninja/examples/powerpoint, demonstrating how you can customize the VDO.Ninja controller or embed the VDO.Ninja controller into your app.

You can pass a room name via the URL, if you wish to use it in production
```
Client link: https://vdo.ninja/examples/powerpoint?room=TESTROOM123
Host link: https://vdo.ninja/?room=TESTROOM123&midiin
```

![image](https://user-images.githubusercontent.com/2575698/210515432-7d1dd7a1-6f68-4f5d-a779-0ec982bc768e.png)


For developers, the IFRAME API commands are `{nextSlide:true}` and  `{prevSlide:true}`. Pretty simple, and this will auto-transmit the commands to any remotely connected peer with `&midiin` added to their URL. If they have the autohotkey script running, and PowerPoint running, it should give you control of their presentation.

ie: `iframe.contentWindow.postMessage({nextSlide:true}, '*');`

This `nextSlide` and `prevSlide` API call is currently only available on VDO.Ninja version 22.12 and up. Refer to the VDO.Ninja IFRAME documentation here https://docs.vdo.ninja/guides/iframe-api-documentation 

If you want to make custom commands, you can just issue RAW MIDI commands via the VDO.Ninja IFRAME API instead.

`iframe.contentWindow.postMessage({"sendRawMIDI":{data:[176, 110, 12]}}, '*');` is an example of a custom MIDI command. You'll need to update the AutoHotKey script to listen for those new MIDI commands, and do something with it, but that's pretty easy in most cases. Adding a button for such a custom command to the sample code provided is just basic javascript/html, which is also quite accessible.

Sending custom raw MIDI commands is supported I think in v20 of VDO.Ninja; not just v22.12 and newer.  As such, if `{nextSlide:true}` isn't supported in your version of VDO.Ninja, you can just send `"sendRawMIDI":{data:[176, 110, 11]}}` for example instead.  `[176, 110, 10]` is for "prevSlide", etc.

## HTTP / Websocket API

Lastly, VDO.Ninja has an HTTP / WSS API, and it lets you issue commands to the host system.  HTTP commands work with StreamDecks, so you can have a remote client issue commands to your VDO.Ninja host system, controlling your PowerPoint, without even that client having VDO.Ninja open themselves.  The reason we'd be using VDO.Ninja is just to make use of the MIDI code that's already there.  

We'd just need to have `https://vdo.ninja/alpha/?api=YOURAPIKEY&midiin` open to make use of this as a host- we don't need to be even viewing or sharing a video.  The API key can be made up, so long as it matches the HTTP request.  This requires VDO.Ninja version 22.12 or newer.

`https://api.vdo.ninja/YOURAPIKEY/nextSlide` and `https://api.vdo.ninja/YOURAPIKEY/prevSlide` woudl be the HTTP commands. Refer to the API documentation for details on the WSS commands.

You can refer to the HTTP/WSS API documentation for more information: https://github.com/steveseguin/Companion-Ninja


## Troubleshooting

If you get a message pop up from the AutoHotkey script, saying "Slide Next" or something, it means it didn't detect PowerPoint. Instead, it's just letting you know it got a request for next/prev slide, but can't act on it.  If PowerPoint is detected, those messages do not now.

If you are not using English, the issue may be that the PowerPoint window name may be something other than "PowerPoint Slide Show". If this is the case, update the run.ahk file's code so that the window name the code is looking for matches the name of the PowerPoint window's name on your system.

If you are using a Mac/Linux, 🤷🏻‍♂️, I'd appreciate any help debugging and improving support there. 

## TO-DO still
- keyboard short cuts for the next/prev need to be added
- Google Slide support and other app support added
- maybe add some addition buttons in; do I need more than prev/next ?
