Attribute VB_Name = "Module1"
' This module contains all the standard 3d code
' within the assignment (part 1)

Option Explicit

' declaration of types...

' Matrices
Public Type TmatrixType
 matrix(1 To 4, 1 To 4) As Double
End Type

' Homogenous Points
Public Type ThomogenousPoint
 homopoint(1 To 4) As Double
End Type

' quads
Public Type Tquad
  quad(1 To 4) As ThomogenousPoint
End Type

' lines - 2 homogenous points
Public Type Tline
 aline(1 To 2) As ThomogenousPoint
End Type

' slices - up to 20 lines
Public Type Tslice
 slice(1 To 20) As Tline
End Type

' an array of slices...
' this one is the base object
Public myslice() As Tslice

' the same object as slices but as quads instead...
Public myquads() As Tquad

' this one gets buggered about with.
Public tmpobject() As Tquad

' all the matrices themselves
Public rotationMatrixX As TmatrixType
Public rotationMatrixY As TmatrixType
Public rotationMatrixZ As TmatrixType
Public perspectiveMatrix As TmatrixType
Public translationMatrix As TmatrixType
Public scalingMatrix As TmatrixType




' Pi - used to convert degrees to radians for the rotation algorithms
Public Const Pi As Double = 3.14592653589793

' has the first sweep been done already
Public sweptflag As Boolean

' do I want to apply the changes I've just made to the original object
Public Apply As Boolean

' which transformation do I want
Public drawoption As String

' which scroll bar(axis) has been changed
Public barchanged As String

' Lx and Ly are stored versions of the previous X and Y values
Public Lx As Double
Public Ly As Double

' converted X,Y and Z values
Public Myx As Double
Public Myy As Double
Public Myz As Double

' the number of lines in a particular slice
Public Linecount As Integer

' probably doesn't need to be Public any more, deg = degrees used for rotations
Public deg As Double

' flag used to detect whether this is the start or end of a line
Public Counttmp As Boolean
Public Sub my3dinit()
    Linecount = 0
    sweptflag = False
    Counttmp = False
    drawoption = ""
    barchanged = ""
    Lx = 0
    Ly = 0
    
    Myx = 0
    Myy = 0
    Myz = 0
    ReDim myslice(1 To 30)
End Sub

Public Sub initialisePerspectiveMatrix(f As Double)
    ' similar to the cube program
    Dim row, col As Integer
    For row = 1 To 4
      For col = 1 To 4
        perspectiveMatrix.matrix(row, col) = 0
      Next col
    Next row
    perspectiveMatrix.matrix(1, 1) = 1
    perspectiveMatrix.matrix(2, 2) = 1
    perspectiveMatrix.matrix(3, 3) = 1
    perspectiveMatrix.matrix(4, 4) = 1 / f
End Sub

Public Sub initialiseRotationMatrixX(d As Double)
    ' similar to the cube program
    Dim row, col As Integer
    For row = 1 To 4
      For col = 1 To 4
        rotationMatrixX.matrix(row, col) = 0
      Next col
    Next row
    rotationMatrixX.matrix(1, 1) = 1
    rotationMatrixX.matrix(2, 2) = Cos(d)
    rotationMatrixX.matrix(2, 3) = Sin(d)
    rotationMatrixX.matrix(3, 2) = -Sin(d)
    rotationMatrixX.matrix(3, 3) = Cos(d)
    rotationMatrixX.matrix(4, 4) = 1
End Sub
Public Sub initialiseRotationMatrixY(d As Double)
    ' similar to the cube program
    Dim row, col As Integer
    For row = 1 To 4
      For col = 1 To 4
        rotationMatrixY.matrix(row, col) = 0
      Next col
    Next row
    rotationMatrixY.matrix(1, 1) = Cos(d)
    rotationMatrixY.matrix(1, 3) = -Sin(d)
    rotationMatrixY.matrix(2, 2) = 1
    rotationMatrixY.matrix(3, 1) = Sin(d)
    rotationMatrixY.matrix(3, 3) = Cos(d)
    rotationMatrixY.matrix(4, 4) = 1
End Sub
Public Sub initialiseRotationMatrixZ(d As Double)
    ' similar to the cube program
    Dim row, col As Integer
    For row = 1 To 4
      For col = 1 To 4
        rotationMatrixZ.matrix(row, col) = 0
      Next col
    Next row
    rotationMatrixZ.matrix(1, 1) = Cos(d)
    rotationMatrixZ.matrix(1, 2) = Sin(d)
    rotationMatrixZ.matrix(2, 1) = -Sin(d)
    rotationMatrixZ.matrix(2, 2) = Cos(d)
    rotationMatrixZ.matrix(3, 3) = 1
    rotationMatrixZ.matrix(4, 4) = 1
End Sub
Public Sub initialiseTranslationMatrix(tx, ty, tz As Double)
    ' similar to the cube program
    Dim row, col As Integer
    
    For row = 1 To 4
      For col = 1 To 4
        translationMatrix.matrix(row, col) = 0
      Next col
    Next row
    translationMatrix.matrix(1, 1) = 1
    translationMatrix.matrix(2, 2) = 1
    translationMatrix.matrix(3, 3) = 1
    translationMatrix.matrix(4, 4) = 1
    translationMatrix.matrix(1, 4) = tx
    translationMatrix.matrix(2, 4) = ty
    translationMatrix.matrix(3, 4) = tz
End Sub

Public Sub initialiseScalingMatrix(Sx As Double, Sy As Double, Sz As Double)
    ' similar to the cube program
    Dim row, col As Integer
    
    For row = 1 To 4
      For col = 1 To 4
        scalingMatrix.matrix(row, col) = 0
      Next col
    Next row
    scalingMatrix.matrix(1, 1) = Sx
    scalingMatrix.matrix(2, 2) = Sy
    scalingMatrix.matrix(3, 3) = Sz
    scalingMatrix.matrix(4, 4) = 1
End Sub


Public Sub dothesweep()
    ' takes a line and sweeps it, then
    ' stores the results as an object made of quads
    ' and also as an object made of lines
    ReDim myquads(1 To 30 * Linecount)
    ReDim tmpobject(1 To 30 * Linecount) As Tquad
    Dim s, a, c, l As Integer
    
    ' create a rotation object of 30 slices...
    deg = 12 * (Pi / 180)
    initialiseRotationMatrixY (deg)
    For c = 1 To 30
    If c < 30 Then
    myslice(c + 1) = (rotateslicearoundy(myslice(c)))
    End If
    myslice(c) = (dorescaleslice(myslice(c)))
    Next c
    
    ' build the object from quads based on the slices above...
    a = 0
    For c = 1 To 30
        For s = 1 To Linecount
            a = a + 1
            For l = 1 To 4
              myquads(a).quad(1).homopoint(l) = myslice(c).slice(s).aline(1).homopoint(l)
              myquads(a).quad(2).homopoint(l) = myslice(c).slice(s).aline(2).homopoint(l)
                If c < 30 Then
                  myquads(a).quad(3).homopoint(l) = myslice(c + 1).slice(s).aline(2).homopoint(l)
                  myquads(a).quad(4).homopoint(l) = myslice(c + 1).slice(s).aline(1).homopoint(l)
                Else
                  myquads(a).quad(3).homopoint(l) = myslice(1).slice(s).aline(2).homopoint(l)
                  myquads(a).quad(4).homopoint(l) = myslice(1).slice(s).aline(1).homopoint(l)
                End If
            Next l
            tmpobject(a) = myquads(a)
        Next s
    Next c
    
End Sub
Public Sub dothelot()
    ' calls the transformation functions
    ReDim tmpobject(1 To 30 * Linecount) As Tquad
    Dim c, s, l, h As Integer
    
    For c = 1 To (30 * Linecount)
        tmpobject(c) = myquads(c)
        If drawoption = "Rotate" Then
            If barchanged = "X" Then
                tmpobject(c) = (rotatearoundx(myquads(c)))
            End If
            
            If barchanged = "Y" Then
                tmpobject(c) = (rotatearoundy(myquads(c)))
            End If
            
            If barchanged = "Z" Then
                tmpobject(c) = (rotatearoundz(myquads(c)))
            End If
        End If
        
        If drawoption = "Scale" Then
            tmpobject(c) = (doscale(myquads(c)))
        End If
        
        If drawoption = "Translate" Then
            tmpobject(c) = (dotranslate(myquads(c)))
        End If
        
        ' store the last action if applying
        If Apply = True Then
            For s = 1 To Linecount
                For l = 1 To 4
                    For h = 1 To 4
                        myquads(c).quad(l).homopoint(h) = tmpobject(c).quad(l).homopoint(h)
                    Next h
                Next l
            Next s
        End If
        
    Next c
End Sub

Public Function dorescaleslice(tmpslice As Tslice) As Tslice
    ' rescale object slices
    Dim s, l, i As Integer
    
    For s = 1 To Linecount
        For l = 1 To 2
            For i = 1 To 4
                dorescaleslice.slice(s).aline(l).homopoint(i) = tmpslice.slice(s).aline(l).homopoint(i) / tmpslice.slice(s).aline(l).homopoint(4)
            Next i
        Next l
    Next s
End Function

Public Function rotateslicearoundy(tmpslice As Tslice) As Tslice
    ' used for the initial sweep, returns slices, not quads
    Dim row, col, s, l As Integer
    Dim total As Single
    ' Rotation around Y bits
    For s = 1 To Linecount
        For l = 1 To 2
            For row = 1 To 4
                total = 0
                For col = 1 To 4
                   total = total + tmpslice.slice(s).aline(l).homopoint(col) * rotationMatrixY.matrix(row, col)
                Next col
                rotateslicearoundy.slice(s).aline(l).homopoint(row) = total
            Next row
        Next l
    Next s
End Function


Public Function rotatearoundx(tmpquad As Tquad) As Tquad
    ' rotation around the X axis - returns quads
    Dim row, col, l As Integer
    Dim total As Single
    ' Rotation around x bits
    For l = 1 To 4
        For row = 1 To 4
            total = 0
            For col = 1 To 4
               total = total + tmpquad.quad(l).homopoint(col) * rotationMatrixX.matrix(row, col)
            Next col
            rotatearoundx.quad(l).homopoint(row) = total
        Next row
    Next l
End Function

Public Function rotatearoundy(tmpquad As Tquad) As Tquad
    ' rotation around the Y axis - returns quads
    Dim row, col, l As Integer
    Dim total As Single
    
    ' Rotation around Y bits
    For l = 1 To 4
        For row = 1 To 4
            total = 0
            For col = 1 To 4
               total = total + tmpquad.quad(l).homopoint(col) * rotationMatrixY.matrix(row, col)
            Next col
            rotatearoundy.quad(l).homopoint(row) = total
        Next row
    Next l
End Function

Public Function rotatearoundz(tmpquad As Tquad) As Tquad
    ' rotation around the Z axis - returns quads
    Dim row, col, l As Integer
    Dim total As Single
    
    ' Rotation around z bits
    For l = 1 To 4
        For row = 1 To 4
            total = 0
            For col = 1 To 4
               total = total + tmpquad.quad(l).homopoint(col) * rotationMatrixZ.matrix(row, col)
            Next col
            rotatearoundz.quad(l).homopoint(row) = total
        Next row
    Next l
End Function

Public Function doscale(tmpquad As Tquad) As Tquad
    ' scaling transformations - returns quads
    Dim row, col, l As Integer
    Dim total As Single
    
    ' scaling along various axes...
    For l = 1 To 4
        For row = 1 To 4
            total = 0
            For col = 1 To 4
               total = total + tmpquad.quad(l).homopoint(col) * scalingMatrix.matrix(row, col)
            Next col
            doscale.quad(l).homopoint(row) = total
        Next row
    Next l
End Function

Public Function dotranslate(tmpquad As Tquad) As Tquad
    ' translation transformations - returns quads
    Dim row, col, l As Integer
    Dim total As Single
    
    ' moving along various axes...
    For l = 1 To 4
        For row = 1 To 4
            total = 0
            For col = 1 To 4
               total = total + tmpquad.quad(l).homopoint(col) * translationMatrix.matrix(row, col)
            Next col
            dotranslate.quad(l).homopoint(row) = total
        Next row
    Next l
End Function
Public Function dorescalebits(tmpquad As Tquad) As Tquad
    ' rescale homogenous points - returns quads
    Dim l, i As Integer
    
    For l = 1 To 4
        For i = 1 To 4
            dorescalebits.quad(l).homopoint(i) = tmpquad.quad(l).homopoint(i) / tmpquad.quad(l).homopoint(4)
        Next i
    Next l
End Function

Public Function doperspectivebits(tmpquad As Tquad) As Tquad
    ' perspective projection
    Dim row, col, l As Integer
    Dim total As Single
    
    ' - Perspective bits
    For l = 1 To 4
        For row = 1 To 4
            total = 0
            For col = 1 To 4
               total = total + tmpquad.quad(l).homopoint(col) * perspectiveMatrix.matrix(row, col)
            Next col
            doperspectivebits.quad(l).homopoint(row) = total
        Next row
    Next l

End Function



Public Function dofittowindow(tmpquad As Tquad) As Tquad
    ' a fiddle to display the object in the middle of my VB picture window
    Dim l As Integer
     
    For l = 1 To 4
        dofittowindow.quad(l).homopoint(1) = tmpquad.quad(l).homopoint(1) + 3450
        dofittowindow.quad(l).homopoint(2) = -tmpquad.quad(l).homopoint(2) + 3810
        dofittowindow.quad(l).homopoint(3) = tmpquad.quad(l).homopoint(3)
        dofittowindow.quad(l).homopoint(4) = tmpquad.quad(l).homopoint(4)
    Next l

End Function
