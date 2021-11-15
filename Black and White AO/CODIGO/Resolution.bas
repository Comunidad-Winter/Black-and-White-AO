Attribute VB_Name = "Resolution"
Public Sub SetResolucion()
 
    Dim lRes As Long
    Dim MidevM As typDevMODE
    
    lRes = EnumDisplaySettings(0, 0, MidevM)
    
    Dim intWidth As Integer
    Dim intHeight As Integer
    
    oldResWidth = Screen.width \ Screen.TwipsPerPixelX
    oldResHeight = Screen.height \ Screen.TwipsPerPixelY
    
    Dim CambiarResolucion As Boolean
    If NoRes Then
        CambiarResolucion = (oldResWidth < 800 Or oldResHeight < 600)
    Else
        CambiarResolucion = (oldResWidth <> 800 Or oldResHeight <> 600)
    End If
    
    If CambiarResolucion Then
        
         With MidevM
               .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
               .dmPelsWidth = 800
               .dmPelsHeight = 600
               .dmBitsPerPel = 16
           End With
       Else
          With MidevM
               .dmFields = DM_BITSPERPEL
             .dmBitsPerPel = 16
          End With
      End If
         If lRes = ChangeDisplaySettings(MidevM, CDS_TEST) Then
     End If
End Sub
 
Public Sub ResetResolucion()
 
    Dim typDevM As typDevMODE
    Dim lRes As Long
    
    If Not bNoResChange Then
    
        lRes = EnumDisplaySettings(0, ENUM_CURRENT_SETTINGS, typDevM)
        
        With typDevM
            .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
            .dmPelsWidth = oldResWidth
            .dmPelsHeight = oldResHeight
            .dmBitsPerPel = oldDepth
            .dmDisplayFrequency = oldFrequency
        End With
        
        lRes = ChangeDisplaySettings(typDevM, CDS_TEST)
    End If
End Sub


