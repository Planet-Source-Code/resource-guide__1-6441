VERSION 5.00
Begin VB.Form frmRes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "RESOURCE FILE DEMO"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadIt 
      Caption         =   "CLICK TO LOAD PICTURES"
      Height          =   390
      Left            =   660
      TabIndex        =   0
      Top             =   135
      Width           =   3675
   End
   Begin VB.Image Image2 
      Height          =   930
      Left            =   1950
      Stretch         =   -1  'True
      Top             =   4020
      Width           =   1230
   End
   Begin VB.Image Image1 
      Height          =   2925
      Left            =   225
      Stretch         =   -1  'True
      Top             =   780
      Width           =   4785
   End
End
Attribute VB_Name = "frmRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'###########################################################
'# All code copyright Andy McCurtin 2000                   #
'# This guide is all my own work please don't nick it      #
'# If you want to distribute it or post is on your web     #
'# site you can all I ask is that you e-mail me so that I  #
'# can visit the site.                                     #
'# If you have questions or comments please tell me        #
'#          E-mail : andy_mccurtin@yahoo.com               #
'###########################################################


'Before you read on
'------------------
'This is a quick guide to using resource files in your VB
'projects.  First you need to find RC.exe in you VB directory
'(Usually under \tools or \Wizards you may also need Rcdll.dll)
'When you have found these files copy them to a directory
'thats easier to find (If you can't find them e-mail me and
'I'll send you them in a *.zip file).

'The Beginning
'-------------
'Now you've got RC.exe & Rcdll.dll, you must make an *.rc file
'you can do this through notepad or any other editor (But as
'I'm a great guy I've included a test file that you can rename
'and edit to suit your needs).  When you have your *.rc file
'you need to describe the files you want to include in the
'resource file, you also need to include a reference for the
'file.  The systax is as follows :-
'       nameID keyword filename

'What the above syntax means
'---------------------------
'nameID is the reference of the file you can call it whatever
'you want, however using a descriptive name halps when coding
'i.e. A picture of a bomb namID = Bomb etc.

'keyword is the type of file you are referening to, this can
'be any of the following :-
'Bitmap, Cursor, Icon, Sound or Video (Sound is a *.Wav file &
'Video is a *.AVI file)

'filename is the full path of the file you wish to include in
'the resource i.e. D:\Images\Head.bmp etc.

'Example
'-------
'For this project I used a Bitmap called Head.bmp and an Icon
'called Note.ico so to include them in the resource I used the
'following lines in the *.rc file
'               Head BITMAP D:\Head.bmp
'               Note ICON D:\Note.ico

'You can keep adding to you *.rc file until you have included
'all the files you need

'End Game
'--------
'Once you've gathered all the files you need into your *.rc
'file you need to compile it into a Resource file(*.res)
'First save you *.rc file then load up the DOS window
'You need to be in the directory you have RC.exe in
'If you've never used DOS before (Shame on you) here's a
'VERY brief guide. If you want to access another drive type
'(Drive letter):
'e.g. C:
'If you want to access a certain directory type cd(Directory)
'e.g. cd images
'To access subdirectories use the same command

'When you are in the correct directory here's what you type :-

'RC/r filename.rc

'filename.rc is the name of you *.rc file
'If sucessful DOS will go back to the prompt
'You should now have a *.res file

'A good tip I use is to store my *.rc files in the same directory
'as RC.exe, it's easier to keep track of them that way and it's
'easier to compile them.

'Using the newly made *.res file in VB
'-------------------------------------
'Load VB press CTRL + D to add a file, once you've added the
'*.res file to your project you need to access it.
'LoadResPicture is used to load Bitmap's, Icon's and Cusor's.
'LoadResData is used to load Wav's and AVI's (Data loaded
'using LoadResData can no bigger than 64K)

'To load a Bitmap into an Image use the following :-
'Image1.Picture = LoadResPicture("Reference",0)
'Change Reference to the reference you used in the *.rc file
'the 0 means that we are loading a Bitmap, when loading
'Icons you use 1 and for Cursors use 2
'Using LoadResData returns a string containing the actual Bits
'in the resource.

'NOTE !!!
'--------
'When you have your *.res file you need to keep it, by this I
'mean that when you compile you VB program you will notice that
'you don't need the *.res file to run it, however if you want
'to edit the source code you will need the *.res file.
'I know this sounds obvious but I've made this mistake my self
'and it can be really annoying

'As it's now 1am on a Sunday morning, I'm tired starting to
'see double and have way too much coffee in my system I'm
'not going to go into the details of how to use LoadResData
'the help files can tell you all about that (that is what
'their there for, isn't it?)
'Explained in the code below is how to use LoadResPicture.
'If I get up early enough today (Doubtful!!) I may add more
'on how to use LoadResData.

'That's it I hope this guide has been helpful to you
'If it hasn't sorry I tried my best.
'Enjoy the code anyway
'                       ANDY


Private Sub cmdLoadIt_Click()
'You can also load pictures into PictureBoxes and DirectDraw
'surfaces.
    Image1.Picture = LoadResPicture("Head", 0)
    Image2.Picture = LoadResPicture("Note", 1)
    
    Me.Icon = LoadResPicture("note", 1)
    Me.Caption = "<<<<<<<<< All hail Andy and his amazing talent !!!!"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MsgBox "Thanks for downloading my guide" & vbCr & _
        "Any problems e-mail : andy_mccurtin@yahoo.com", vbInformation
End Sub
