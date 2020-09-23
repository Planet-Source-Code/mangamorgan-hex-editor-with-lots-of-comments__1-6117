VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Text            =   "C:\WINDOWS\CALC.EXE"
      Top             =   2640
      Width           =   3555
   End
   Begin VB.Label Label1 
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Because ive used option explicit, every variable has to
'be declared before it is used. All the variables used are
'declared below, with its perpose!

Dim z As Long
' z is used as a FOR...LOOP variable
Dim n As Long
' n is the number of blocks required
Dim b As Integer
' b is another FOR...LOOP variable.
Dim byt As Byte
' this is basically used to store a single byte of data!
' Vital for getting each HEX value!
Dim h As String * 2
' This is where the hex value will be stored. Because it is
' always going to be 2 characters long ive forced the computer
' to ONLY make that much space. It speed the process up because
' less memory is being used!
Dim block() As String * 256
' This is where the initial dump of the file will be placed. Its 256
' long because i found it to be large enough to store a great deal
' of data, but small enough to not take up too much memory. This
' ratio i found give the best performance, but u can experement for
' your self AS LONG AS you change every 256 nuber to your new number!
' It is an array because it saves memory space and again speeds up the
' process.
Dim hexOutput() As String
' this is where 1 long line of the hex values produced will be stored. It is
' an array and for the same reasons as above!
Dim finalOutput As String
' The whole output will be dumped in here before being dumped in the rich
' textbox. This is done because it is quicker!
Dim file As String
' this is the file in which you are going to open.


Private Sub Form_Click()
    file = Text1.Text
    ' get the file name of the file you wish to get a hex read out on.
    Print "Reading in file (" & file & ")..."
    ' for the GUI effect only!
    Open file For Binary As #1
    ' the file needs to be opened as a binary file because we need a
    ' binary read out of it!
        n = FileLen(file) / 256
        ' this is calculating the number of blocks ("n") required to store
        ' the file in by getting the file size & deviding it by the size of
        ' the blocks.
        ReDim block(n)
        ' This is resizing the array of blocks to snugly fit the size of the file.
        ' We do this because it saves the amount of memory required to store all
        ' thus leaving more memory to run its opperations thus speeding up the
        ' program!
        z = 0
        ' z is the number of blocks that have been read. We need to set it to
        ' zero an manualy incriment it by 1 because we are using a DO...LOOP
        ' instead of a FOR...LOOP. I did this because we need to knnow when
        ' the program has reached the end of the file (EOF)
        Do Until EOF(1) = True
            Get #1, , block(z)
            ' this is collecting a block of data at a time and storing each
            ' block of data in a different index in the array.
            z = z + 1
            ' a manual incriment!
        Loop
    Close #1
    ' we dont need to deal with the file on disk any more! no that it is in the
    ' memory we can manipulate it much, much, much fater!!
    Print "File loaded into memory! " & n & " blocks used!"
    Print "Compiling HEX output..."
    ' 2 nice little GUI status messages! very helpful! :)
    ReDim hexOutput(n)
    ' i am now resizing the array for the lines of hex data the program will
    ' produce soon. This is done in the same way and for the same reasons as
    ' when we did it previously ("ReDim block(n)") - NOTE: both arrays are the
    ' same size (the same amount of indexs)!
    For z = 0 To n
    ' loop the following process the amount of times that there are indexs in
    ' the resized arrays - ie the ammount of blocks that where required!
        For b = 1 To 256
        ' now loop the amount of times that the block size is!
            byt = Asc(Mid(block(z), b, 1))
            h = Hex(byt)
            ' this bit is quite cleaver (if I do say so myself!) Basically it
            ' gets each character, in turn, from the string (the block of data)
            ' and gets its character code number (eg the ASCII code) which is
            ' always 1 byte big! That is the exact size we need to get the 2
            ' digit hex number (by using the HEX function). it is then stored in
            ' a 2 digit long string.
            If Right(h, 1) = " " Then h = "0" & Left(h, 1)
            ' this bit checks to see if the number was a single digit or not.
            ' If it was then it sticks a 0 at front thus giving the standard
            ' form of displaying hex numbers (eg 2F 3E 02 00 instead of 2F 3E 2 0)
            hexOutput(z) = hexOutput(z) & " " & h
            ' now it is adding the new found hex byte to all of the older hex
            ' bytes.
        Next b
        Label1.Caption = "Block " & z & " of " & n & " completed..."
        Label1.Refresh
        ' More GUI related stuff! This time its is kind of a timer gauge!
    Next z
    Print
    Print "Colating data..."
    ' Yet more GUI!
    DoEvents
    ' needed for GUI! ;) Fixes a refresh bug in VB!
    For z = 0 To n
        finalOutput = finalOutput & hexOutput(z)
    Next z
    ' loops around the number of times that there is blocks of data and adds
    ' each block of hex onto the end of the last. (all of the above hex bytes
    ' where stored in blocks to for speed and consistancy!)
    Print "Triming overflow byte values..."
    DoEvents
    ' GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI! GUI!
    Do Until Right(finalOutput, 2) <> "00"
        finalOutput = Left(finalOutput, Len(finalOutput) - 3)
    Loop
    ' This last loop basically keeps looping until the last few hex byts are NOT
    ' 00! This is done because 1 bug in my program meeds that if the file is not
    ' exactly a multiple of the size of a block, then it will add blanks ("00")
    ' to round it up to 1. (u cant use half blocks - well, not in this version
    ' anyway!) So this is kind of a bug fix - abet a slow 1 :( - It is not required
    ' but it makes the final output that much more professional!
    finalOutput = LTrim(finalOutput)
    ' This cleans up any previewing spaces. It is needed, but it is ultrafast!
    Form2.txtHex.Text = finalOutput
    ' This puts the whole of the generated hex stuff in the rich textbox. An
    ' ordanery textbox just can't hold enough text!!
    Print "HEX dump of " & file & " created!"
    ' guess what??? GUI!
    Form2.Show
    ' show the hex stuff in a new window!
End Sub
