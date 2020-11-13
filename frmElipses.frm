VERSION 5.00
Begin VB.Form frmElipses 
   Caption         =   "Elipses"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmElipses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Me.AutoRedraw = True
  'Me.Circle (3000, 3000), 1000, vbBlue, 1, 1, 0.6

  For a = 1 To 2000
    Me.Circle (3000, 1000), 1000, vbBlack, 1, 1, 0.6
    Me.Circle (3000, 1000 + a), 1000, vbBlue, 1, 1, 0.6
    Me.Circle (3000, 3000), 1000, vbBlack, 1, 1, 0.6
  Next a

  For b = 1 To 2000
    Me.Circle (6000, 1000), 1000, vbBlack, 1, 1, 0.6
    Me.Circle (6000, 1000 + b), 1000, vbGreen, 1, 1, 0.6
    Me.Circle (6000, 3000), 1000, vbBlack, 1, 1, 0.6
  Next b

  For c = 1 To 1000
    Me.Circle (10000, 1000), 500, vbBlack, 1, 1, 0.6
    Me.Circle (10000, 1000 + c), 500 + c, vbCyan, 1, 1, 0.6
    Me.Circle (10000, 2000), 1500, vbBlack, 1, 1, 0.6
  Next c

  'aro
  For d = 1 To 1000
    Me.Circle (13500, 3000), 1500, vbBlack, 1, 1, 0.6
    Me.Circle (13500, 3000 + d), 500 + c, vbRed, 1, 1, 0.6
    Me.Circle (13500, 4000), 1500, vbBlack, 1, 1, 0.6
  Next d

  'cono
  For e = 1 To 1000
    Me.Circle (3000, 6000 - e * 2), 1000 - e, vbYellow, 1, 1, 0.6
    Me.Circle (3000, 6000), 1000, vbBlack, 1, 1, 0.6
  Next e

  'esfera
  For f = 1 To 450
    f = 2 * f
    f1 = f
    Me.Circle (6000, 6000 - f), 1000 - f, , 1, 1, 0.6
    Me.Circle (6000, 6000), 1000, vbBlack, 1, 1, 0.6
    Me.Circle (6000, 6000 + f1), 1000 - f1, , 1, 1, 0.6
  Next f

  'figura rara
  'For e = 1 To 1000
  'pi = 3.14

  'Me.Circle (3000, 5200), 250, vbBlack, 1, 1, 0.6
  'Me.Circle (3000, 5200 + e), 250 * e, vbBlack, 1, 1, 0.6
  'Me.Circle (3000, 6000), 1000, vbBlack, 1, 1, 0.6
  'Next e

  'Figura rara
  'For e = 1 To 1000
  'pi = 3.14

  'Me.Circle (3000, 5200), 250, vbBlack, 1, 1, 0.6
  'Me.Circle (3000, 5200 + e), 250 * e, vbBlack, 1, 1, 0.6
  'Me.Circle (3000, 6000), 1000, vbBlack, 1, 1, 0.6
  'Next e
End Sub
