Sub bam_gio()
    Dim ims As Integer
    Dim igiay As Integer
    Dim imin As Integer
    
    ims = 0
    igiay = 0
    imin = 0
    
    Do
        Range("a3").Value = imin
        Range("b3").Value = igiay
        Range("c3").Value = ims
        ims = ims + 1
        If ims = 100 Then
            ims = 0
            igiay = igiay + 1
            If igiay = 60 Then
                igiay = 0
                imin = imin + 1
            End If
        End If
     Loop Until imin = 10
End Sub
