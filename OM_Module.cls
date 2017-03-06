VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OM_Module"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'***************************************************************************'
'                                                                           '
'                                                                           '
'   CLASS:     OM_Module                                                    '
'                                                                           '
'   AUTHOR(S): Theodore J. Rosenthal                                        '
'              Jeff P. Chrstos                                              '
'                                                                           '
'   DATE:      January 23, 2013                                             '
'                                                                           '
'   COPYRIGHT:                                                              '
'                                                                           '
'   Copyright (c) by Systems Technology Incorporated:                       '
'                                                                           '
'   ALL RIGHTS RESERVED. The data and methods contained in this document    '
'   are proprietary to Systems Technology Incorporated.                     '
'                                                                           '
'   DESCRIPTION:                                                            '
'                                                                           '
'   Class containing all of the methods that the STISIM Drive Open Module   '
'   capability supports.                                                    '
'                                                                           '
'   Methods:                                                                '
'                                                                           '
'              AddNew - Uses the Open Module event in the SDL to initiate   '
'                       a new process                                       '
'       ControlInputs - Allows the user to modify the driver's control      '
'                       inputs                                              '
'            Dynamics - Allows the user to modify the vehicle dynamics      '
'         HandleCrash - Occurs anytime there is a crash                     '
'          Initialize - Handles any initialization that must occur before   '
'                       the simulation starts                               '
'   InitializeTerrain - Initialize the roadway terrain parameters before    '
'                       the simulation starts                               '
'             PostRun - Handle any final issues before the software exits   '
'    SavePlaybackData - Saves data into the playback file that will be      '
'                       needed by your Open Module during a playback        '
'             StartUp - Handle any processes immediately after the software '
'                       starts                                              '
'            Shutdown - Handles any shutdown processes immediately after a  '
'                       simulation run                                      '
'              Update - Allows the process to be updated each time through  '
'                       the simulation loop                                 '
'                                                                           '
'                                                                           '
'***************************************************************************'
'
'
Option Explicit
'
'
'***************************************************************************'
'                                                                           '
'                                                                           '
'   This section contains the definition of all variables that are          '
'   both global and local to this class                                     '
'                                                                           '
'                                                                           '
'***************************************************************************'
'
'
'   Create an instance of the windows tools class
'
'
Dim Tools As New TJRWinToolsCls
'
'
'   Create an instance of the graphics object
'
'
Dim Graphics As New TJR3DGraphics
'
'
'   Create an instance of the terrain object
'
'
Dim Terrain As New STI_3D_Terrain
'
'
'   Define the generic DirectSound sound object
'
'
Dim Ds As DirectSound8
Dim Dx As New DirectX8
'
'
'   Dimension all variables that are public to this class and the calling routine
'
'
Public BSAVData As String
Public ErrorMessage As String
Public LogFileHandle As Long
Public NewForm As Form
Public SaveControls As Integer
Public TextMessage As String
Public WillHandleCrash As Long
'
'
'   Create a type for the driver input information
'
'
Type DriverControlInputs
    Steer As Single                     ' Steering angle count from the controller card
    Throttle As Single                  ' Throttle control count from the controller card
    Brake As Single                     ' Brake control count from the controller card
    Gear As Integer                     ' Current transmission gear
    Buttons As Long                     ' Current state of the driver's input buttons
End Type
Dim Driver As DriverControlInputs
'
'
'   Define all vairables that will be global to this class
'
'
Dim Bool As Boolean                     ' Temporary boolean variable for use anytime a throw away boolean is needed
Dim DynVars As OMDynamicVariables       ' UDT containing STISIM Drive variables that change as the run progresses
Dim PlaybackMode As Integer             ' Flag specifying if it is real mode (0) or playback (1)
Dim St As String                        ' Temporary string variable for use anytime a throw away string is needed
Dim StaticVars As OMStaticVariables     ' UDT containing STISIM Drive variables that are fixed by the simulator
Dim WorldIndex() As Long                ' Index for the main graphics world
'
'
Type vehicle
    Acceleration As Single                  ' Vehicle acceleration
    BrakeModel As Long                      ' Brake light model handle
    index As Long                           ' Model index
    InitialHeading As Single                ' Vehicle initial heading when it is activated
    Lat As Double                           ' Lane postion with respect to the roadway centerline
    Lon As Double                           ' Distance into the run (from the beginning)
    SixDOF As SixDOFPosition                ' Inertial orientation of the vehicle
    Speed As Single                         ' Vehicle speed
    SpinDuration As Single                  ' Time it takes to complete on wheel rotation
    SpinModel As Long                       ' Spin animation model handle
    SpinSpeed As Single                     ' Vehicle tire rotation speed
    VisFlag As Long                         ' Visibility flag
End Type
Dim v() As vehicle



Dim VAhead As SurrVehicle
Dim VNextAhead As SurrVehicle
Dim LAhead As SurrVehicle
Dim RAhead As SurrVehicle
Dim Lbehind As SurrVehicle
Dim Rbehind As SurrVehicle
Dim Traffic As TRFQuery



Public Function Update(dv As OMDynamicVariables, vehicle As DYNAMICSParams, Events As SimEvents, NumEvents As Integer, EDist() As Single, EDes() As Integer, EIndex() As Integer, SimSounds() As SoundEffects) As Boolean
'
'
'   Function for handling all Open Module action during the actual simulation loop
'
'
'   Parameters:
'
'            DV - User defined type containing the simulation parameters that are changing at each time step
'       Vehicle - User defined type containing the driver's vehicle dynamic variables
'        Events - UDT that contains the parameter settings for each supported event
'     NumEvents - Number of events that are in the current display list
'       EDist() - Distance from the driver to the event
'        EDes() - Event designator for each active event
'      EIndex() - Event index for each event in the display list. This value is the index into the Events UDT
'                 so that you can get the parameters for each individual event in the display list
'   SimSounds() - Contains the sound buffer references to the simulator sound effects
'


'This calls the traffic querying update function that queries the traffic and updates the position of the traffic objects.
' it takes the same arguments as the Update function of the OM_Module with an additional 'Max Range' parameter that determines
' how far ahead it considers traffic (currently set to 500ft)
Call Traffic.QueryTraffic(vehicle, dv, NumEvents, EDes(), EIndex(), Events, 1000)

If Traffic.getFirstAhead(VAhead) Then
    dv.DisplayStrings(1) = CStr(VAhead.range)
Else
    dv.DisplayStrings(1) = "No vehicle detected"
End If
If Traffic.getLeftAhead(LAhead) Then
    dv.DisplayStrings(2) = CStr(LAhead.range)
Else
    dv.DisplayStrings(2) = "No vehicle detected"
End If
If Traffic.getRightAhead(RAhead) Then
    dv.DisplayStrings(3) = CStr(RAhead.range)
Else
    dv.DisplayStrings(3) = "No vehicle detected"
End If
If Traffic.getLeftBehind(Lbehind) Then
    dv.DisplayStrings(4) = CStr(Lbehind.range)
Else
    dv.DisplayStrings(4) = "No vehicle detected"
End If
If Traffic.getRightBehind(Rbehind) Then
    dv.DisplayStrings(5) = CStr(Rbehind.range)
Else
    dv.DisplayStrings(5) = "No vehicle detected"
End If

'
'
On Error GoTo ErrorOccurred
'
'
'   Make the dynamic variables available to all other methods
'
'
DynVars = dv
'
'
'   Setup the return from function
'
'
Update = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Update")
Update = False
'
'
End Function

'***************************************************************************'
'                                                                           '
'                                                                           '
'   This section contains the public methods that the main STISIM Drive     '
'   modules can access during the course of a simulation run                '
'                                                                           '
'                                                                           '
'***************************************************************************'
'
'
Public Function AddNew(OMVars As OMParameters) As Boolean
'
'
'   Function for adding a new interactive Open Module event
'
'
'   Parameters:
'
'   OMVars - User defined type containing the parameters for the given Open Module being acted on
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
AddNew = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("AddNew")
AddNew = False
'
'
End Function

Public Function ControlInputs(Dyn As DYNAMICSParams, Steering As Single, Throttle As Single, Brake As Single, Gear As Integer, DInput As Long) As Boolean
'
'
'   Function for handling any user defined control inputs
'
'
'   Parameters:
'
'        Dyn - User defined type containing simulation dynamics variables
'   Steering - Steering wheel angle input digital count
'   Throttle - Throttle pedal input digital count
'      Brake - Braking pedal input digital count
'       Gear - Current transmission gear
'     DInput - Current button values
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Make the driver inputs available to the other methods
'
'
With Driver
    .Brake = Brake
    .Steer = Steering
    .Throttle = Throttle
    .Buttons = DInput
End With
'
'
'   Setup the return from function
'
'
ControlInputs = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("ControlInputs")
ControlInputs = False
'
'
End Function

Public Function Dynamics(Dyn As DYNAMICSParams) As Boolean
'
'
'   Function for handling all Open Module dynamic updates
'
'
'   Parameters:
'
'   Dyn - User defined type containing the driver's vehicle dynamic variables
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
Dynamics = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Dynamics")
Dynamics = False
'
'
End Function

Public Function HandleCrash(Override As Integer, Events As SimEvents, CrashEvent As Integer, EventIndex As Integer) As Boolean
'
'
'   Function for handling all Open Module action if there is a crash during the simulation run
'
'
'   Parameters:
'
'      Override - Parameter defining how STISIM Drive will handle the crash when this
'                 method returns control to it
'        Events - UDT that contains the parameter settings for each supported event
'    CrashEvent - Event designator for the event that caused the crash
'    EventIndex - Index specifying which instance of the crash event caused the crash
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
Override = 0
HandleCrash = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("HandleCrash")
HandleCrash = False
'
'
End Function

Public Function Initialize(SV As OMStaticVariables, ParamFile As String, WorldIndexIn() As Long, GraphicsIn As TJR3DGraphics) As Boolean
'
'
'   Function for handling all Open Module initialization
'
'
'   Parameters:
'
'             SV - User defined type containing simulation static variables
'      ParamFile - Name of a file that contains any parameters that will be required
'                  by the Open Module code
'   WorldIndex() - Handle for the various graphics context that hold the roadway environments
'     GraphicsIn - Reference to the graphics object that the main simulator uses to draw the 3D world
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Dimension all variables local to this routine
'
'
Dim FileNum As Integer
Dim i As Integer
Dim NumVerts As Integer


Set Traffic = New TRFQuery

'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Make the static variables available to all other methods
'
'
StaticVars = SV
'
'
'   Assign references to the main simulator's graphics object and the world indices so that they can be used throughout the OM
'
'
Set Graphics = GraphicsIn
NumVerts = UBound(WorldIndexIn)
ReDim WorldIndex(NumVerts)
For i = 0 To NumVerts
    WorldIndex(i) = WorldIndexIn(i)
Next
'
'
'   Only act if it is the center display system
'
'
If StaticVars.DisplaySystem = "CenterDisplay" Then
    '
    '
    '   Setup any labels that will be used to display data in the STISIM Drive runtime window display
    '
    '
    With SV
        .DisplayStrings(1) = "Lead Vehicle Distance"
        .DisplayStrings(2) = "Left Side Lead vehicle"
        .DisplayStrings(3) = "Right Side Lead Vehicle"
        .DisplayStrings(4) = "Rear Left"
        .DisplayStrings(5) = "Rear Right"
    End With
    '
    '
    '   If there is an initialization file specified then do the initializing
    '
    '
    If Len(ParamFile) Then
        '
        '
        '   Handle the initial parameter file
        '
        '
        FileNum = FreeFile(0)
        Open ParamFile For Input As #FileNum
        Close FileNum
    End If
End If
'
'
'   Setup the return from function
'
'
Initialize = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Initialize")
Initialize = False
'
'
End Function

Public Function InitializeTerrain(TerrainIn As STI_3D_Terrain) As Boolean
'
'
'   Function for setting the terrain parameters
'
'
'   Parameters:
'
'   TerrainIn - Reference to the terrain object that is used by the main simulation loop
'
'   Returns:
'
'   True if the terrain data was setup correctly, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Initialize the terrain object
'
'
Set Terrain = TerrainIn
'
'
'   Setup the return from function
'
'
InitializeTerrain = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("InitializeTerrain")
InitializeTerrain = False
'
'
End Function

Public Function PostRun(Comments As String, DriverName As String, RunNumber As String, DriverID As String) As Boolean
'
'
'   Function for handling anything before the software exists
'
'
'   Parameters:
'
'     Comments - Comments entered in the subject information form
'   DriverName - Name of the driver from the subject information form
'    RunNumber - Run number entered in the subject information form
'     DriverID - ID entered from the subject information form
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Release the objects that were created
'
'
Set Ds = Nothing
Set Dx = Nothing
Set Tools = Nothing
'
'
'   Setup the return from function
'
'
PostRun = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("PostRun")
PostRun = False
'
'
End Function

Public Function SavePlaybackData(PlaybackData() As Single, PlaybackString As String) As Boolean
'
'
'   Function for specifying any OM data that will be stored as part of a playback file
'
'
'   Parameters:
'
'     PlaybackData - Array containing the data that will be saved
'   PlaybackString - String containing string data that will be saved
'
'   Returns:
'
'   True if everything worked fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the return from function
'
'
SavePlaybackData = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("SavePlaybackData")
SavePlaybackData = False
'
'
End Function

Public Function Shutdown(RunCompleted As Integer) As Boolean
'
'
'   Function for handling Open Module processes immediately after a simulation run has ended
'
'
'   Parameters:
'
'   RunCompleted - Flag specifying if the run completed successfully or not
'
'                    0 - Aborted before start of run
'                    1 - Run completed successfully
'                  > 1 - Aborted during the run
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Release the objects that were created
'
'
Set Graphics = Nothing
Set Terrain = Nothing
'
'
'   Setup the return from function
'
'
Shutdown = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("Shutdown")
Shutdown = False
'
'
End Function

Public Function StartUp(Config As GAINSParams, BackForm As Object, SV As OMStaticVariables, UseNew As Boolean, PlaybackData() As Single, PlaybackString As String) As Boolean
'
'
'   Function for handling Open Module processes immediately after the software is started
'
'
'   Parameters:
'
'           Config - Configuration file parameters
'         BackForm - Current STISIM Drive back ground form
'               SV - User defined type containing simulation static variables
'           UseNew - Flag specifying if a new background form will be used (True) or not (False)
'   PlaybackData() - Array containing any data that is being transfered from the playback file back into your module
'   PlaybackString - String containing any string data that is being transfered from the playback file back into your module
'
'   Returns:
'
'   True if everything initialized fine, otherwise false. If false use the ErrorMessage
'   parameter to return a message that the program can display to the user
'
'
'
'
'   Set the error handling
'
'
On Error GoTo ErrorOccurred
'
'
'   Setup the sound object
'
'
If SV.SoundOn = True Then
    Set Ds = Dx.DirectSoundCreate(vbNullString)
    Call Ds.SetCooperativeLevel(SV.ActiveHandle, DSSCL_NORMAL)
End If
'
'
'   Assign the type of mode (driving, or playback)
'
'
PlaybackMode = Config.PlaybackMode
'
'
'   Setup the return from function
'
'
UseNew = False
StartUp = True
Exit Function
'
'
'   Handle any errors
'
'
ErrorOccurred:
'
'
ErrorMessage = ProcessError("StartUp")
StartUp = False
'
'
End Function


'
'
'**************************************************************************'
'                                                                          '
'                                                                          '
'   This section contains private routines that are need by and only       '
'   available to the methods in this module                                '
'                                                                          '
'                                                                          '
'**************************************************************************'

Private Function ProcessError(ModuleName As String) As String
'
'
'   Routine for adding information to the error message that will be returned when the
'   Open Module encounters problems
'
'   Parameters:
'
'   ModuleName - Name of the method where the error occured
'
'   Returns:
'
'   New error message including number and description and other information
'
'
'
'
'   Build the error message
'
'
St = "Simulation run aborted! An error has occurred in Open Module " & ModuleName & ":" & vbCrLf & vbCrLf
St = St & "Error number:" & vbTab & Trim(Str(Err.Number)) & vbCrLf
St = St & "Description:" & vbTab & Err.Description & vbCrLf
St = St & "Error source:" & vbTab & Err.Source & vbCrLf
ProcessError = St & "Last DLL Error:" & vbTab & Err.LastDllError & vbCrLf
Bool = Tools.WriteToTJRFile(StaticVars.LogFileHandle, ProcessError)
'
'
End Function

