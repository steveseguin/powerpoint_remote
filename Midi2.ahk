;
; Midi2.ahk for AutoHotkey v2
; Add MIDI input event handling to your AutoHotkey scripts
; https://github.com/hetima/AutoHotkey-Midi
;
; Based on Midi.ahk by Danny Warren
; https://github.com/dannywarren/AutoHotkey-Midi
;

Persistent

; Midi class interface
Class AHKMidi
{

  ; delegate
  static _delegate := False
  delegate
  {
    get => AHKMidi._delegate
    set => AHKMidi._delegate := value
  }

  ; Enable or disable label event handling
  static midiLabelCallbacks := True

  ; Enable specific process callback like explorer_exe_MidiNoteOnC(event){}
  static _specificProcessCallback := False
  specificProcessCallback
  {
    get => AHKMidi._specificProcessCallback
    set => AHKMidi._specificProcessCallback := value
  }

  ; Enable or disable path through event that not handled to output device
  static _midiEventPassThrough := False
  midiEventPassThrough
  {
    get => AHKMidi._midiEventPassThrough
    set => AHKMidi._midiEventPassThrough := value
  }
  ; use SetPassThroughDeviceName(deviceName)
  ; if not set, pass trough to all opened midi out devices
  static passThroughDeviceHandle  := False


  _settingFilePath := False
  settingFilePath
  {
    get => this._settingFilePath
    set {
      this._settingFilePath := value
      this.LoadIOSetting(this._settingFilePath)
    }
  }
  ; Defines the string size of midi devices returned by windows (see mmsystem.h)
  static MIDI_DEVICE_NAME_LENGTH := 32

  ; Defines the size of a midi input struct MIDIINCAPS (see mmsystem.h)
  static MIDI_DEVICE_IN_STRUCT_LENGTH  := 44

  ; Defines the size of a midi input struct MIDIINCAPS (see mmsystem.h)
  static MIDI_DEVICE_OUT_STRUCT_LENGTH  := 50

  ; Defines for midi event callbacks (see mmsystem.h)
  static MIDI_CALLBACK_WINDOW   := 0x10000
  static MIDI_CALLBACK_TASK     := 0x20000
  static MIDI_CALLBACK_FUNCTION := 0x30000

  ; Defines for midi event types (see mmsystem.h)
  static MIDI_OPEN      := 0x3C1
  static MIDI_CLOSE     := 0x3C2
  static MIDI_DATA      := 0x3C3
  static MIDI_LONGDATA  := 0x3C4
  static MIDI_ERROR     := 0x3C5
  static MIDI_LONGERROR := 0x3C6
  static MIDI_MOREDATA  := 0x3CC

  ; Defines the size of the standard chromatic scale
  static MIDI_NOTE_SIZE := 12

  ; Defines the midi notes 
  static MIDI_NOTES     := [ "C", "Cs", "D", "Ds", "E", "F", "Fs", "G", "Gs", "A", "As", "B" ]

  ; Defines the octaves for midi notes
  static MIDI_OCTAVES   := [ -2, -1, 0, 1, 2, 3, 4, 5, 6, 7, 8 ]

  ; This is where we will keep the most recent midi in event data so that it can
  ; be accessed via the Midi object, since we cannot store it in the object due
  ; to how events work
  ; We will store the last event by the handle used to open the midi device, so
  ; at least we won't clobber midi events from other devices if the user wants 
  ; to fetch them specifically
  static __midiInEvent        := {}
  ;;;;static __midiInHandleEvent  := {}

  ; List of all midi input/output devices on the system
  static __midiInDevices := {}
  static __midiOutDevices := {}

  ; Default label prefix
  static midiLabelPrefix := "Midi"

  ; Holds a refence to the system wide midi dll, so we don't have to open it
  ; multiple times
  static __midiDll := 0

  ; Always use gui mode when using the midi library, since we need something to
  ; attach midi events to
  ; Gui, +LastFound
  ; Gui("+LastFound") ?
  ; The window to attach the midi callback listener to, which will default to
  ; our gui window
  static __midiInCallbackWindow := Gui()

  ; List of midi input/output devices to listen to messages for, we do this globally
  ; since only one instance of the class can listen to a device anyhow
  static __midiInOpenHandles := Map()
  static __midiOutOpenHandles := Map()

  ; Count of open handles, since ahk doesn't have a method to actually count the
  ; members of an array (it instead just returns the highest index, which isn't
  ; the same thing)
  static __midiInOpenHandlesCount := 0
  static __midiOutOpenHandlesCount := 0

  ; Enable or disable lazy midi in event debugging via tooltips
  static midiEventTooltips  := False

  ; Instance creation
  __New()
  {
    this._settingFilePath := False
    ; global midiDelegate
    this.__MidInDevicesMenu := Menu()
    this.__MidOutDevicesMenu := Menu()

    ; Initialize midi environment
    this.LoadMidi()
    this.QueryMidiInDevices()
    this.QueryMidiOutDevices()
    this.SetupDeviceMenus()

  }

  ; Instance destruction
  __Delete()
  {

    ; Close all midi in/out devices and then unload the midi environment
    this.CloseMidiIns()
    this.CloseMidiOuts()
    this.UnloadMidi()

  }


  ; Load midi dlls
  LoadMidi()
  {
    If ( ! AHKMidi.__midiDll ){
      AHKMidi.__midiDll := DllCall( "LoadLibrary", "Str", "winmm.dll", "Ptr" )
      If ( ! AHKMidi.__midiDll )
      {
        MsgBox("Missing system midi library winmm.dll")
        ExitApp
      }
    }
  }


  ; Unload midi dlls
  UnloadMidi()
  {
    If ( AHKMidi.__midiDll )
    {
      DllCall( "FreeLibrary", "Ptr", AHKMidi.__midiDll )
    }

  }


  ; Open midi in device and start listening
  OpenMidiIn( midiInDeviceId )
  {

    this.__OpenMidiIn( midiInDeviceId )
  
  }

  ; Open midi in device and start listening
  OpenMidiInByName( midiInDeviceName )
  {
    For key, value In AHKMidi.__midiInDevices{
      if (value.deviceName==midiInDeviceName){
        this.__OpenMidiIn( key )
        return key
      }
    }
    return -1
  }


  ; Close midi in device and stop listening
  CloseMidiIn( midiInDeviceId )
  {

    this.__CLoseMidiIn( midiInDeviceId )
    
  }


  ; Close all currently open midi in devices
  CloseMidiIns()
  {

    If ( ! AHKMidi.__midiInOpenHandlesCount )
    {
      Return
    }

    ; We have to store the handles we are going to close in advance, because
    ; autohotkey gets confused if we are removing things from an array while
    ; iterative over it
    deviceIdsToClose := {}

    ; Iterate once to get a list of ids to close
    For midiInDeviceId In AHKMidi.__midiInOpenHandles
    {
      deviceIdsToClose.Push( midiInDeviceId )
    }

    ; Iterate again to actually close them
    For index, midiInDeviceId In deviceIdsToClose
    {
      this.CloseMidiIn( midiInDeviceId )
    }

  }


  ; Query the system for a list of active midi input devices
  QueryMidiInDevices()
  {
    midiInDevices := Map()

    deviceCount := DllCall( "winmm.dll\midiOutGetNumDevs" ) - 1

    Loop deviceCount 
    {

      midiInDevice := {}

      deviceNumber := A_Index - 1

      midiInStruct := Buffer( AHKMidi.MIDI_DEVICE_IN_STRUCT_LENGTH, 0 )

      midiQueryResult := DllCall( "winmm.dll\midiInGetDevCapsA", "UINT", deviceNumber, "PTR", midiInStruct, "UINT", AHKMidi.MIDI_DEVICE_IN_STRUCT_LENGTH )

      ; Error handling
      If ( midiQueryResult )
      {
        MsgBox("Failed to query midi devices")
        Return
      }

      manufacturerId := NumGet( midiInStruct, 0, "USHORT" )
      productId      := NumGet( midiInStruct, 2, "USHORT" )
      driverVersion  := NumGet( midiInStruct, 4, "UINT" )
      deviceName     := StrGet( midiInStruct.Ptr + 8, AHKMidi.MIDI_DEVICE_NAME_LENGTH, "CP0" )
      ; support        := NumGet( midiInStruct, 4, "UINT" )

      midiInDevice.direction      := "IN"
      midiInDevice.deviceNumber   := deviceNumber
      midiInDevice.deviceName     := deviceName
      midiInDevice.productID      := productID
      midiInDevice.manufacturerID := manufacturerID
      midiInDevice.driverVersion  := ( driverVersion & 0xF0 ) . "." . ( driverVersion & 0x0F )

      __MidiEventDebug( midiInDevice )

      midiInDevices[deviceNumber] := midiInDevice
    }

    AHKMidi.__midiInDevices := midiInDevices

  }



  ; Open midi in device and start listening
  OpenMidiOut( midiOutDeviceId )
  {

    this.__OpenMidiOut( midiOutDeviceId )
  
  }
   ; Open midi out device and start listening
  OpenMidiOutByName( midiOutDeviceName )
  {

    For key, value In AHKMidi.__midiOutDevices{
      if (value.deviceName==midiOutDeviceName){
        this.__OpenMidiOut( key )
        return key
      }
    }
    return -1
  }


  ; Close midi in device and stop listening
  CloseMidiOut( midiOutDeviceId )
  {

    this.__CLoseMidiOut( midiOutDeviceId )
    
  }


  ; Close all currently open midi in devices
  CloseMidiOuts()
  {

    If ( ! AHKMidi.__midiOutOpenHandlesCount )
    {
      Return
    }

    ; We have to store the handles we are going to close in advance, because
    ; autohotkey gets confused if we are removing things from an array while
    ; iterative over it
    deviceIdsToClose := {}

    ; Iterate once to get a list of ids to close
    For midiOutDeviceId In AHKMidi.__midiOutOpenHandles
    {
      deviceIdsToClose.Push( midiOutDeviceId )
    }

    ; Iterate again to actually close them
    For index, midiOutDeviceId In deviceIdsToClose
    {
      this.CloseMidiOut( midiOutDeviceId )
    }

  }




  ; Query the system for a list of active midi Output devices
  QueryMidiOutDevices()
  {
    midiOutDevices := Map()

    deviceCount := DllCall( "winmm.dll\midiOutGetNumDevs" ) 

    Loop deviceCount 
    {

      midiOutDevice := {}

      deviceNumber := A_Index - 1
      midiOutStruct := Buffer(AHKMidi.MIDI_DEVICE_OUT_STRUCT_LENGTH, 0)

      midiQueryResult := DllCall( "winmm.dll\midiOutGetDevCapsA", "UINT", deviceNumber, "PTR", midiOutStruct, "UINT", AHKMidi.MIDI_DEVICE_OUT_STRUCT_LENGTH )

      ; Error handling
      If ( midiQueryResult )
      {
        MsgBox("Failed to query midi devices")
        Return
      }

      manufacturerId := NumGet( midiOutStruct, 0, "USHORT" )
      productId      := NumGet( midiOutStruct, 2, "USHORT" )
      driverVersion  := NumGet( midiOutStruct, 4, "UINT" )
      deviceName     := StrGet( midiOutStruct.Ptr + 8, AHKMidi.MIDI_DEVICE_NAME_LENGTH, "CP0" )
      ; technology     := NumGet( midiOutStruct, AHKMidi.MIDI_DEVICE_NAME_LENGTH + 8 + 0, "USHORT" )
      ; voices         := NumGet( midiOutStruct, AHKMidi.MIDI_DEVICE_NAME_LENGTH + 8 + 2, "USHORT" )
      ; notes          := NumGet( midiOutStruct, AHKMidi.MIDI_DEVICE_NAME_LENGTH + 8 + 4, "USHORT" )
      ; channelmask    := NumGet( midiOutStruct, AHKMidi.MIDI_DEVICE_NAME_LENGTH + 8 + 8, "USHORT" ) ; +6?
      ; support        := NumGet( midiOutStruct, AHKMidi.MIDI_DEVICE_NAME_LENGTH + 8 + 10, "USHORT" ) ; +8?

      midiOutDevice.direction      := "OUT"
      midiOutDevice.deviceNumber   := deviceNumber
      midiOutDevice.deviceName     := deviceName
      midiOutDevice.productID      := productID
      midiOutDevice.manufacturerID := manufacturerID
      midiOutDevice.driverVersion  := ( driverVersion & 0xF0 ) . "." . ( driverVersion & 0x0F )

      __MidiEventDebug( midiOutDevice )

      midiOutDevices[deviceNumber] := midiOutDevice

    }

    AHKMidi.__midiOutDevices := midiOutDevices

  }

  ; Set up device selection menus
  SetupDeviceMenus()
  {
    __SelectMidiInDevice(ItemName, ItemPos, MyMenu){
      midiInDeviceId := ItemPos-1
      if ( AHKMidi.__midiInOpenHandles.Get(midiInDeviceId, 0) > 0 )
      {
        this.__CloseMidiIn( midiInDeviceId )
      }
      else
      {
        this.__OpenMidiIn( midiInDeviceId )        
      }
      this.SaveIOSetting(this._settingFilePath)
    }
    __SelectMidiOutDevice(ItemName, ItemPos, MyMenu){
      midiOutDeviceId := ItemPos-1
      if ( AHKMidi.__midiOutOpenHandles.Get(midiOutDeviceId, 0) > 0 )
      {
        this.__CloseMidiOut( midiOutDeviceId )
      }
      else
      {
        this.__OpenMidiOut( midiOutDeviceId )        
      }
      this.SaveIOSetting(this._settingFilePath)
    }

    haveInDevices := false
    For key, value In AHKMidi.__midiInDevices
    {
      this.__MidInDevicesMenu.Add(value.deviceName, __SelectMidiInDevice)
      haveInDevices := true
    }
    haveOutDevices := false
    For key, value In AHKMidi.__midiOutDevices
    {
      this.__MidOutDevicesMenu.Add(value.deviceName, __SelectMidiOutDevice)
      haveOutDevices := true
    }
    ;A_TrayMenu.Add()
    ; Menu, Tray, Add
    A_TrayMenu.Add()
    if (haveInDevices==true) {
      A_TrayMenu.Add("MIDI Input  Devices", this.__MidInDevicesMenu)
    }
    if (haveOutDevices==true) {
      A_TrayMenu.Add("MIDI Output Devices", this.__MidOutDevicesMenu)
    }
  }


  ; Returns the last midi in event values
  MidiIn()
  {
    Return AHKMidi.__MidiInEvent
  }

  MidiOutRawData(rawData, deviceHandle := False)
  {
    ; handle
    If (deviceHandle)
    {
      result := DllCall("winmm.dll\midiOutShortMsg", "UInt", deviceHandle, "UInt", rawData, "UInt")
      if (result)
      {
        msgbox("There was an error sending the midi event")
      }
    }else{
      For midiOutDeviceId In AHKMidi.__midiOutOpenHandles
      {
        ;Call api function to send midi event  
        result := DllCall("winmm.dll\midiOutShortMsg"
                  , "UInt", AHKMidi.__midiOutOpenHandles[midiOutDeviceId]
                  , "UInt", rawData
                  , "UInt")
    
        if (result)
        {
          msgbox("There was an error sending the midi event")
        }
      }
    }
  }

  MidiOut(EventType, Channel, Param1, Param2, deviceHandle := False)
  {
    ;handle is handle to midi output device returned by midiOutOpen function
    ;EventType and Channel are combined to create the MidiStatus byte.  
    ;MidiStatus message table can be found at http://www.harmony-central.com/MIDI/Doc/table1.html
    ;Possible values for EventTypes are NoteOn (N1), NoteOff (N0), CC, PolyAT (PA), ChanAT (AT), PChange (PC), Wheel (W) - vals in () are optional shorthand 
    ;SysEx not supported by the midiOutShortMsg call
    ;Param3 should be 0 for PChange, ChanAT, or Wheel.  When sending Wheel events, put the entire Wheel value
    ;in Param2 - the function will split it into it's two bytes 
    ;returns 0 if successful, -1 if not.  
  
    ;Calc MidiStatus byte
    If (EventType = "NoteOn" OR EventType = "N1")
      MidiStatus :=  143 + Channel
    Else if (EventType = "NoteOff" OR EventType = "N0")
      MidiStatus := 127 + Channel
    Else if (EventType = "ControlChange" OR EventType = "CC")
      MidiStatus := 175 + Channel
    Else if (EventType = "PolyAT" OR EventType = "PA")
      MidiStatus := 159 + Channel
    Else if (EventType = "ChanAT" OR EventType = "AT")
      MidiStatus := 207 + Channel
    Else if (EventType = "ProgramChange" OR EventType = "PC")
      MidiStatus := 191 + Channel
    Else if (EventType = "PitchWheel" OR EventType = "PW")
    {  
      MidiStatus := 223 + Channel
      Param2 := Param1 >> 8 ;MSB of wheel value
      Param1 := Param1 & 0x00FF ;strip MSB, leave LSB only
    }
    else
    {
      Return
    }

    ;Midi message Dword is made up of Midi Status in lowest byte, then 1st parameter, then 2nd parameter.  Highest byte is always 0
    dwMidi := MidiStatus + (Param1 << 8) + (Param2 << 16)

    this.MidiOutRawData(dwMidi, deviceHandle)
  }

  DeviceHandleForName(deviceName)
  {
    If (deviceName != "")
    {
      For key, value In AHKMidi.__midiOutDevices{
        if (value.deviceName == deviceName){
          return AHKMidi.__midiOutOpenHandles[value.deviceNumber]
        }
      }
    }
    return False
  }

  MidiOutToDeviceId(EventType, Channel, Param1, Param2, deviceId)
  {
    deviceHandle := AHKMidi.__midiOutOpenHandles[deviceId]
    if(deviceHandle){
      this.MidiOut(EventType, Channel, Param1, Param2, deviceHandle)
    }else{
      msgbox("device " deviceId " not found or not open")
    }
  }

  MidiOutToDeviceName(EventType, Channel, Param1, Param2, deviceName)
  {
    deviceHandle := this.DeviceHandleForName(deviceName)
    if(deviceHandle){
      this.MidiOut(EventType, Channel, Param1, Param2, deviceHandle)
    }else{
      msgbox(deviceName " not found or not open")
    }
  }

  MidiOutRawDataToDeviceId(rawData, deviceId)
  {
    deviceHandle := AHKMidi.__midiOutOpenHandles[deviceId]
    if(deviceHandle){
      this.MidiOutRawData(rawData, deviceHandle)
    }else{
      msgbox("device " deviceId " not found or not open")
    }
  }

  MidiOutRawDataToDeviceName(rawData, deviceName)
  {
    deviceHandle := this.DeviceHandleForName(deviceName)
    if(deviceHandle){
      this.MidiOutRawData(rawData, deviceHandle)
    }else{
      msgbox(deviceName " not found or not open")
    }
  }

  ; save and load I/O setting to ini file
  ; format is like this
  ;   [AutoHotkeyMidi]
  ;   inputDevices=input1name////input2name////
  ;   outputDevices=output1name////
  LoadIOSetting(filePath)
  {
    If (filePath = False or filePath = "")
    {
      Return
    }
    If (not FileExist(filePath)){
      Return
    }
    Try
    {
      setting := IniRead(filePath, "AutoHotkeyMidi", "inputDevices")
      If (setting == "ERROR"){
          Return
      }
      names := StrSplit(setting, "////")
      For index, inDeviceName In names
      {
        If (inDeviceName != ""){
          this.OpenMidiInByName(inDeviceName)
        }
      }
    }
    catch
    {
    }

    Try
    {
      setting := IniRead(filePath, "AutoHotkeyMidi", "outputDevices")
      If (setting == "ERROR"){
          Return
      }
      names := StrSplit(setting, "////")
      For index, outDeviceName In names
      {
        If (outDeviceName != ""){
          this.OpenMidiOutByName(outDeviceName)
        }
      }
    }
    catch
    {
    }
  }

  ;should call in OnExit
  SaveIOSetting(filePath)
  {
    If (filePath = False or filePath = "")
    {
      Return
    }
    Try
    {
      setting := ""
      If (AHKMidi.__midiInOpenHandlesCount != 0)
      {
        For midiInDeviceId In AHKMidi.__midiInOpenHandles
        {
          setting .= AHKMidi.__midiInDevices[midiInDeviceId].deviceName . "////"
        }
      }
      IniWrite setting, filePath, "AutoHotkeyMidi", "inputDevices"

      setting := ""
      If (AHKMidi.__midiOutOpenHandlesCount != 0)
      {
        For midiOutDeviceId In AHKMidi.__midiOutOpenHandles
        {
          setting .= AHKMidi.__midiOutDevices[midiOutDeviceId].deviceName . "////"
        }
      }
      IniWrite setting, filePath, "AutoHotkeyMidi", "outputDevices"
    }
    catch
    {
    }
  }

  SetPassThroughDeviceName(deviceName)
  {
    AHKMidi.passThroughDeviceHandle := this.DeviceHandleForName(deviceName)
  }




  ; Open a handle to a midi device and start listening for messages
  __OpenMidiIn( midiInDeviceId )
  {
    ;;;;global __MidiInHandleEvent

    ; Look this device up in our device list
    device := AHKMidi.__midiInDevices[midiInDeviceId]

    ; Create variable to store the handle the dll open will give us
    ; NOTE: Creating variables this way doesn't work with class variables, so
    ; we have to create it locally and then store it in the class later after
    midiInHandle := Buffer( 4, 0 )

    ; Open the midi device and attach event callbacks
    midiInOpenResult := DllCall( "winmm.dll\midiInOpen", "Ptr", midiInHandle, "UINT", midiInDeviceId, "UINT", AHKMidi.__midiInCallbackWindow.Hwnd, "UINT", 0, "UINT", AHKMidi.MIDI_CALLBACK_WINDOW, "UINT" )

    ; Error handling
    If ( midiInOpenResult || ! midiInHandle )
    {
      MsgBox("Failed to open midi in device")
      Return
    }

    ; Fetch the actual handle value from the pointer
    midiInHandle := NumGet( midiInHandle, "UINT" )

    ; Start monitoring midi signals
    midiInStartResult := DllCall( "winmm.dll\midiInStart", "UINT", midiInHandle, "UINT" )

    ; Error handling
    If ( midiInStartResult )
    {
      MsgBox("Failed to start midi in device")
      Return
    }

    ; Create a spot in our global event storage for this midi input handle
    ;;;;__MidiInHandleEvent[midiInHandle] := {}

    ; Register a callback for each midi event
    ; We only need to do this once for all devices, so only do it if we are
    ; the first device to be opened
    if ( AHKMidi.__midiInOpenHandlesCount = 0)
    {
      OnMessage( AHKMidi.MIDI_OPEN,      __MidiInCallback )
      OnMessage( AHKMidi.MIDI_CLOSE,     __MidiInCallback )
      OnMessage( AHKMidi.MIDI_DATA,      __MidiInCallback )
      OnMessage( AHKMidi.MIDI_LONGDATA,  __MidiInCallback )
      OnMessage( AHKMidi.MIDI_ERROR,     __MidiInCallback )
      OnMessage( AHKMidi.MIDI_LONGERROR, __MidiInCallback )
      OnMessage( AHKMidi.MIDI_MOREDATA,  __MidiInCallback ) 
    }

    ; Add this device handle to our list of open devices
    AHKMidi.__midiInOpenHandles[midiInDeviceId] := midiInHandle

    ; Increase the tally for the number of open handles we have
    AHKMidi.__midiInOpenHandlesCount++

    ; Check this device as enabled in the menu
    this.__MidInDevicesMenu.Check(device.deviceName)

  }


  __CloseMidiIn( midiInDeviceId )
  {
    ; Look this device up in our device list
    device := AHKMidi.__midiInDevices[midiInDeviceId]

    ; Unregister callbacks if we are the last open handle
    if ( AHKMidi.__midiInOpenHandlesCount <= 1 )
    {
      OnMessage( AHKMidi.MIDI_OPEN,      __MidiInCallback, 0 )
      OnMessage( AHKMidi.MIDI_CLOSE,     __MidiInCallback, 0 )
      OnMessage( AHKMidi.MIDI_DATA,      __MidiInCallback, 0 )
      OnMessage( AHKMidi.MIDI_LONGDATA,  __MidiInCallback, 0 )
      OnMessage( AHKMidi.MIDI_ERROR,     __MidiInCallback, 0 )
      OnMessage( AHKMidi.MIDI_LONGERROR, __MidiInCallback, 0 )
      OnMessage( AHKMidi.MIDI_MOREDATA,  __MidiInCallback, 0 )
    }

    ; Destroy any midi in events that might be left over
    ;;;;__MidiInHandleEvent[midiInHandle] := {}

    ; Stop monitoring midi
    midiInStopResult := DllCall( "winmm.dll\midiInStop", "UINT", AHKMidi.__midiInOpenHandles[midiInDeviceId], "UINT" )

    ; Error handling
    If ( midiInStopResult != 0)
    {
      MsgBox("Failed to stop midi in device")
      Return
    }

    ; Close the midi handle
    midiInStopResult := DllCall( "winmm.dll\midiInClose", "UINT", AHKMidi.__midiInOpenHandles[midiInDeviceId], "UINT" )

    ; Error handling
    If ( midiInStopResult != 0)
    {
      MsgBox("Failed to close midi in device")
      Return
    }

    ; Finally, remove the handle from the array
    AHKMidi.__midiInOpenHandles.Delete( midiInDeviceId )

    ; Decrease the tally for the number of open handles we have
    AHKMidi.__midiInOpenHandlesCount--

    ; Uncheck this device in the menu
    this.__MidInDevicesMenu.Uncheck(device.deviceName)

  }



  ; Open a handle to a midi device and start listening for messages
  __OpenMidiOut( midiOutDeviceId )
  {
    ; Look this device up in our device list
    device := AHKMidi.__midiOutDevices[midiOutDeviceId]

    ; Create variable to store the handle the dll open will give us
    ; NOTE: Creating variables this way doesn't work with class variables, so
    ; we have to create it locally and then store it in the class later after
    midiOutHandle := Buffer( 4, 0 )
    dwFlags := 0

    ; Open the midi device and attach event callbacks
    midiOutOpenResult := DllCall( "winmm.dll\midiOutOpen"
      , "Ptr", midiOutHandle
      , "UINT", midiOutDeviceId
      , "UINT", 0
      , "UINT", 0
      , "UINT", dwFlags
      , "UInt" )

    ; Error handling
    If ( midiOutOpenResult || ! midiOutHandle )
    {
      MsgBox("Failed to open midi out device")
      Return
    }

    ; Fetch the actual handle value from the pointer
    midiOutHandle := NumGet( midiOutHandle, "UINT" )

    ; Add this device handle to our list of open devices
    AHKMidi.__midiOutOpenHandles[midiOutDeviceId] := midiOutHandle

    ; Increase the tally for the number of open handles we have
    AHKMidi.__midiOutOpenHandlesCount++

    ; Check this device as enabled in the menu
    this.__MidOutDevicesMenu.Check(device.deviceName)

    return

  }


  __ClosemidiOut( midiOutDeviceId )
  {
    ; Look this device up in our device list
    device := AHKMidi.__midiOutDevices[midiOutDeviceId]

    
    ; Close the midi handle
    midiOutStopResult := DllCall( "winmm.dll\midiOutClose", "UINT", AHKMidi.__midiOutOpenHandles[midiOutDeviceId] , "UINT")

    ; Error handling
    If ( midiOutStopResult )
    {
      MsgBox("Failed to close midi in device")
      Return
    }

    ; Finally, remove the handle from the array
    AHKMidi.__midiOutOpenHandles.Delete( midiOutDeviceId )

    ; Decrease the tally for the number of open handles we have
    AHKMidi.__midiOutOpenHandlesCount--

    ; UnCheck this device in the menu
    this.__MidOutDevicesMenu.Uncheck(device.deviceName)

  }

}


; Event callback for midi input event
; Note that since this is a callback method, it has no concept of "this" and
; can't access class members
__MidiInCallback( wParam, lParam, msg, hwnd )
{
  ;;;;global __midiInHandleEvent
  ; Will hold the midi event object we are building for this event
  midiEvent := {}

  ; Will hold the labels we call so the user can capture this midi event, we
  ; always start with a generic ":Midi" label so it always gets called first
  labelCallbacks := Array()

  ; Grab the raw midi bytes
  rawBytes := lParam

  ; Split up the raw midi bytes as per the midi spec
  highByte  := lParam & 0xF0 
  lowByte   := lParam & 0x0F
  data1     := (lParam >> 8) & 0xFF
  data2     := (lParam >> 16) & 0xFF

  ; Determine the friendly name of the midi event based on the status byte
  if ( highByte == 0x80 || ( highByte == 0x90 && data2 == 0 ) )
  {
    midiEvent.status := "NoteOff"
  }
  else if ( highByte == 0x90 )
  {
    midiEvent.status := "NoteOn"
  }
  else if ( highByte == 0xA0 )
  {
    midiEvent.status := "Aftertouch"
  }
  else if ( highByte == 0xB0 )
  {
    midiEvent.status := "ControlChange"
  }
  else if ( highByte == 0xC0 )
  {
    midiEvent.status := "ProgramChange"
  }
  else if ( highByte == 0xD0 )
  {
    midiEvent.status := "ChannelPressure"
  }
  else if ( highByte == 0xE0 )
  {
    midiEvent.status := "PitchWheel"
  }
  else if ( highByte == 0xF0 )
  {
    midiEvent.status := "Sysex"
  }
  else
  {
    Return
  }

  ; Add a label callback for the status, ie ":MidiNoteOn"
  labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.status )

  ; Determine how to handle the one or two data bytes sent along with the event
  ; based on what type of status event was seen
  if ( midiEvent.status == "NoteOff" || midiEvent.status == "NoteOn" || midiEvent.status == "AfterTouch" )
  {

    ; Store the raw note number and velocity data
    midiEvent.noteNumber  := data1
    midiEvent.velocity    := data2

    ; Figure out which chromatic note this note number represents
    noteScaleNumber := Mod( midiEvent.noteNumber, AHKMidi.MIDI_NOTE_SIZE )

    ; Look up the name of the note in the scale
    midiEvent.note := AHKMidi.MIDI_NOTES[ noteScaleNumber + 1 ]

    ; Determine the octave of the note in the scale 
    noteOctaveNumber := Floor( midiEvent.noteNumber / AHKMidi.MIDI_NOTE_SIZE )

    ; Look up the octave for the note
    midiEvent.octave := AHKMidi.MIDI_OCTAVES[ noteOctaveNumber + 1 ]

    ; Create a friendly name for the note and octave, ie: "C4"
    midiEvent.noteName := midiEvent.note . midiEvent.octave

    ; Add label callbacks for notes, ie ":MidiNoteOnA", ":MidiNoteOnA5", ":MidiNoteOn97"
    labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.status . midiEvent.note )
    labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.status . midiEvent.noteName )
    labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.status . midiEvent.noteNumber )

  }
  else if ( midiEvent.status == "ControlChange" )
  {

    ; Store controller number and value change
    midiEvent.controller := data1
    midiEvent.value      := data2

    ; Add label callback for this controller change, ie ":MidiControlChange12"
    labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.status . midiEvent.controller )

  }
  else if ( midiEvent.status == "ProgramChange" )
  {

    ; Store program number change
    midiEvent.program := data1

    ; Add label callback for this program change, ie ":MidiProgramChange2"
    labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.status . midiEvent.program )

  }
  else if ( midiEvent.status == "ChannelPressure" )
  {
    
    ; Store pressure change value
    midiEvent.pressure := data1

  }
  else if ( midiEvent.status == "PitchWheel" )
  {

    ; Store pitchwheel change, which is a combination of both data bytes 
    midiEvent.pitch := ( data2 << 7 ) + data1

  }
  else if ( midiEvent.status == "Sysex" )
  {

    ; Sysex messages have another status byte that indicates which type of sysex
    ; message it is (the high byte, which is normally used for the midi channel,
    ; is used for this instead)
    if ( lowByte == 0x0 )
    {
      midiEvent.sysex := "SysexData"
      midiEvent.data  := data1 ;"byte1" ???
    }
    if ( lowByte == 0x1 )
    {
      midiEvent.sysex := "Timecode"
    }
    if ( lowByte == 0x2 )
    {
      midiEvent.sysex     := "SongPositionPointer"
      midiEvent.position  := ( data2 << 7 ) + data1
    }
    if ( lowByte == 0x3 )
    {
      midiEvent.sysex   := "SongSelect"
      midiEvent.number  := data1
    }
    if ( lowByte == 0x6 )
    {
      midiEvent.sysex := "TuneRequest"
    }
    if ( lowByte == 0x8 )
    {
      midiEvent.sysex := "Clock"
    }
    if ( lowByte == 0x9 )
    {
      midiEvent.sysex := "Tick"
    }
    if ( lowByte == 0xA )
    {
      midiEvent.sysex := "Start"
    }
    if ( lowByte == 0xB )
    {
      midiEvent.sysex := "Continue"
    }
    if ( lowByte == 0xC )
    {
      midiEvent.sysex := "Stop"
    }
    if ( lowByte == 0xE )
    {
      midiEvent.sysex := "ActiveSense"
    }
    if ( lowByte == 0xF )
    {
      midiEvent.sysex := "Reset"
    }
    
    ; Add label callback for sysex event, ie: ":MidiClock" or ":MidiStop"
    labelCallbacks.Push( AHKMidi.midiLabelPrefix . midiEvent.sysex )

  }

  ; Channel is always handled the same way for all midi events except sysex
  if ( midiEvent.status != "Sysex" )
  {
    midiEvent.channel := lowByte + 1
  }

  ; Always include the raw midi data, just in case someone wants it
  midiEvent.rawBytes  := rawBytes
  midiEvent.highByte  := highByte
  midiEvent.lowByte   := lowByte
  midiEvent.data1     := data1
  midiEvent.data2     := data2

  ; Store this midi in event in our global array of midi messages, so that the
  ; appropriate midi class an access it later
  AHKMidi.__MidiInEvent := midiEvent
  ;;;;__MidiInHandleEvent[wParam] := midiEvent

  ; Iterate over all the label callbacks we built during this event and jump
  ; to them now (if they exist elsewhere in the code)
  midiEvent.eventHandled := False

  If ( AHKMidi.midiLabelCallbacks )
  {
    If ( AHKMidi._specificProcessCallback ){
      processPrefix := __GetProcessLabel()
    }else{
      processPrefix := False
    }
    For labelName In labelCallbacks
    {
      If ( processPrefix ){
        processLabel := processPrefix . labelName
        If ( HasMethod(AHKMidi._delegate, processLabel, 1) )
        {
          midiEvent.eventHandled := True
          AHKMidi._delegate.%processLabel%(midiEvent)
        } 
      }
      If (midiEvent.eventHandled = False and HasMethod(AHKMidi._delegate, labelName, 1) )
      {
        midiEvent.eventHandled := True
        AHKMidi._delegate.%labelName%(midiEvent)
      }  
    }
  }

  ; Call debugging if enabled
  __MidiEventDebug( midiEvent )

  ; pass through to midi out
  if ( AHKMidi._midiEventPassThrough && ! midiEvent.eventHandled && AHKMidi.__midiOutOpenHandlesCount > 0 )
  {
    if(AHKMidi.passThroughDeviceHandle){
      midiOutResult := DllCall( "winmm.dll\midiOutShortMsg", "UINT", AHKMidi.passThroughDeviceHandle, "UINT", rawBytes )
    }else{
      for deviceId, hndl In AHKMidi.__midiOutOpenHandles
      {
        midiOutResult := DllCall( "winmm.dll\midiOutShortMsg", "UINT", hndl, "UINT", rawBytes )
      }
    }
  }
}

__GetProcessLabel()
{
  try{
    result := WinGetProcessName("A")
    result := StrReplace(result, "." , "_")
    result := StrReplace(result, A_Space, "_")
    result := StrReplace(result, "#", "_")
    Return result . "_"
  }catch as e{
    ;type(e) = "TargetError"
  }
  Return False
}


; Send event information to a listening debugger
__MidiEventDebug( midiEvent )
{

  ; debugStr := ""

  ; For key, value In midiEvent
  ;   debugStr .= key . ":" . value . "`n"

  ; debugStr .= "---`n"

  ; ; Always output event debug to any listening debugger
  ; OutputDebug debugStr 

  ; ; If lazy tooltip debugging is enabled, do that too
  ; if (AHKMidi.midiEventTooltips){
  ;   ToolTip debugStr
  ; }

}
