; Designed for Version 2 of AutoHotKey
; Requires a Virtual MIDI controller; loopMIDI is recommended

; Based on libraries / code snippets from:	
; https://github.com/hetima/AutoHotkey-Midi
; https://nothans.com/stream-deck-autohotkey-powerpoint

; Written by by Steve Seguin, 2023.
; Designed to allow for remote control of PPT via VDO.Ninja's WebMIDI API.

#include "Midi2.ahk"

midi := AHKMidi()
midi.midiEventPassThrough := True
midi.delegate := MyDelegate()

Class MyDelegate
{
    MidiControlChange(event) {
        ; MsgBox(event.controller . "=" . event.value)
		if (event.controller == 110){ ; setup to work with VDO.Ninja's &ppt mode
			if (event.value==10){
				if (WinExist("PowerPoint Slide Show")){
					WinActivate ;
					Send "{PgUp}"
				}
			}
			if (event.value==11){
				if (WinExist("PowerPoint Slide Show")){
					WinActivate ;
					Send "{PgDn}"
				}
			}
		}
    }
}


MsgBox("MIDI PowerPoint Controller Started")