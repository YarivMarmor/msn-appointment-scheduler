Attribute VB_Name = "Module1"
Public dataArray_S() As String
Public dataArray() As Double
Public vDistribution As Integer
Public vInstances As Long
Public sFileName As String
Public sDistribution As String
Public iNumberOfAppointment As Integer
Public iNumberOfServers As Integer
Public dStartOfDay, dReplications As Double
Public dQoS, dMu, dSD As Double
Public d_mu, d_sd, d_a, d_b As Double
Public dNoShow, dWaitingTime As Double

Public Sub f_AppointmentCalculation()

'Finding the appointment for patients with mu and sd with high quality of service (QoS)
    'Shortcut: Ctrl+Shift+M

    Dim Tij() As Double 'scan duration
    Dim Cij() As Double 'appointment completion time
    Dim Cij1() As Double
    Dim Ai() As Double ' appointment time
    Dim Ai1() As Double 'average appointment start
    Dim Fi1() As Variant
    Dim Fis() As Variant
    Dim Fi() As Variant ' appointment finish
    Dim Di() As Double ' Appointment Delay time
    Dim Si() As Double ' Service Level
    Dim Ti() As Double ' avg Duration
    Dim Mu As Double 'the expected duration time
    Dim S As Double 'the sd of appointment duration
    Dim QoS As Double ' the quality of service needed
    Dim iIndex As Integer
    Dim dummy As Double
    
    Worksheets("Appointment Table").Activate
    Worksheets("Appointment Table").Range("A2:g100000").ClearContents
    formAppointmentClaculation.Show
    
    ReDim Tij(iNumberOfAppointment, vInstances)
    ReDim Cij(iNumberOfAppointment, vInstances)
    ReDim Cij1(iNumberOfAppointment, vInstances)
    ReDim Ai(iNumberOfAppointment + 1)
    ReDim aAi(iNumberOfAppointment + 1)
    ReDim Fi(iNumberOfAppointment + 1, vInstances)
    ReDim Di(iNumberOfAppointment + 1)
    ReDim Si(iNumberOfAppointment + 1)
    ReDim Ti(iNumberOfAppointment + 1)
    ReDim Fi1(vInstances)
    ReDim Fis(vInstances)
    
    'parameters
    'Ai( 1) = dStartOfDay * 60
    QoS = dQoS
    Mu = dMu
    S = d_sd
            
    'replication
    dReplications = 1 '0
    
    For o = 1 To dReplications
    
        For k = 1 To iNumberOfServers
            Ai(k) = dStartOfDay * 60
        Next k
            

        
        'appointment
        For i = 1 To iNumberOfAppointment
            
            Di(i) = 0
            Si(i) = 0
            Ti(i) = 0
            
            'Randomizing the scan duration Tij for each job i and instance j
            For j = 1 To vInstances
            
                
        
                'using a transfromation F function to randomize a normal distribution for positive Tij
                
                Tij(i, j) = -1
                
                'normal distribution
                If vDistribution = 0 Then
                
                    Do While Tij(i, j) < 0
                        'randomize a number from 0 to 1 without the edges from uniform distribution
                        rand_n = Rnd()
                        Do While rand_n = 0 Or rand_n = 1
                            rand_n = Rnd()
                        Loop
                        
                        If Rnd() <= dNoShow Then
                            Tij(i, j) = 0
                        Else: Tij(i, j) = Application.WorksheetFunction.Norm_Inv(rand_n, dMu, dSD)
                        End If
                        
                    Loop
                
                End If
                
                ' uniform distribution
                If vDistribution = 1 Then
                
                    Do While Tij(i, j) < 0
                        'randomize a number from 0 to 1 from uniform distribution
                        rand_n = Rnd()
                        
                        If Rnd() <= dNoShow Then
                            Tij(i, j) = 0
                        Else: Tij(i, j) = dMu + rand_n * (dSD - dMu)  ' dMU is a, dSD is b
                        End If
                        
                    Loop
                
                End If
                
                ' Empirical distribution
                If vDistribution = 2 Then
                
                    Do While Tij(i, j) < 0
                        'randomize a number from 0 to 1
                        rand_n = Rnd()
                        
                        iIndex = WorksheetFunction.Floor_Math(rand_n * (UBound(dataArray) - 1))
                        If Rnd() <= dNoShow Then
                            Tij(i, j) = 0
                        Else: Tij(i, j) = dataArray(iIndex) + (rand_n * (UBound(dataArray) - 1) - iIndex) * (dataArray(iIndex + 1) - dataArray(iIndex))  ' dMU is a, dSD is b
                        End If
                        
                        
                        
                        
                        'Let U = RN (0,1) , A = (n - 1) U, i = trunc (A) + 1; Return X = x(i) + (A - i +1) ( x(i+1) - x(i) )

                    Loop
                
                End If
                               
                Ti(i) = Ti(i) + Tij(i, j) / vInstances ' Ti is the average Tij at appointment i
                
 '               'calculate the end of the scan for the 1st appointment
 '               If i = 1 Then
 '                   Fi(j) = Ai(1) + Tij(i, j)
 '                   Cij(i, j) = Fi(j)
 '               End If
            
            Next j
        
            'Calculating the 2nd appointment per QoS precentile
 '           If i = 1 Then Ai(2) = Application.WorksheetFunction.Percentile_Inc(Fi, QoS)

        Next i

        
        ' order the first Kth Tij
        For j = 1 To vInstances
            For k = 1 To iNumberOfServers - 1
                For kk = k + 1 To iNumberOfServers
                
            
                    If Tij(k, j) > Tij(kk, j) Then
                        dummy = Tij(kk, j)
                        Tij(kk, j) = Tij(k, j)
                        Tij(k, j) = dummy
                    End If
                

                    
                Next kk
                
            Next k
            
            ' 1st will get the next appointment
            Fi(1, j) = Ai(1) + Tij(1, j)
           ' Fi(1, j) = Tij(1, j)
            Fis(j) = Fi(1, j) - dWaitingTime '*** reducing waiting time in order to choose next appointment ***
            
            'updating the end of the first K appoitnment ending (will denote Tij from now)
            For k = 1 To iNumberOfServers
                Tij(k, j) = Ai(k) + Tij(k, j)
            Next k
            
        Next j

        
        'Ai(4) = Ai(1) + Application.WorksheetFunction.Percentile_Inc(Fis, QoS)
        Ai(iNumberOfServers + 1) = Application.WorksheetFunction.Percentile_Inc(Fis, QoS)
        
        ' Calculating the delays in starting time
        For j = 1 To vInstances
        
            If Fis(j) + dWaitingTime > Ai(iNumberOfServers + 1) Then
                Di(iNumberOfServers + 1) = Di(iNumberOfServers + 1) + (Fis(j) + dWaitingTime - Ai(iNumberOfServers + 1)) / vInstances
                Si(iNumberOfServers + 1) = Si(iNumberOfServers + 1) + (1) / vInstances
            End If
            
        Next j
        'finding appoitnemnt by QoS per scan - latter on we will fix it for i>2
  
  
        For scan = iNumberOfServers + 1 To iNumberOfAppointment
            For j = 1 To vInstances
                Tij(scan, j) = Tij(scan, j) + Application.WorksheetFunction.Max(Ai(scan), Tij(scan - iNumberOfServers, j))
            Next j
        
            For j = 1 To vInstances
                
                For k = 0 To iNumberOfServers - 2
                    For kk = k + 1 To iNumberOfServers - 1
                
                        If Tij(scan - kk, j) > Tij(scan - k, j) Then
                            dummy = Tij(scan - kk, j)
                            Tij(scan - kk, j) = Tij(scan - k, j)
                            Tij(scan - k, j) = dummy
                        End If
                        
                    Next kk
                Next k
                
 
                
            ' 1st will get the next appointment
            Fi(scan - (iNumberOfServers - 1), j) = Tij(scan - (iNumberOfServers - 1), j) - dWaitingTime
            Fis(j) = Fi(scan - (iNumberOfServers - 1), j)
        
            Next j


            Ai(scan + 1) = Application.WorksheetFunction.Percentile_Inc(Fis, QoS)
            
                    ' Calculating the delays in starting time
            For j = 1 To vInstances
                If Fis(j) + dWaitingTime > Ai(scan + 1) Then
                    Di(scan + 1) = Di(scan + 1) + (Fis(j) + dWaitingTime - Ai(scan + 1)) / vInstances
                    Si(scan + 1) = Si(scan + 1) + (1) / vInstances
                End If
            Next j
    
        Next scan
    
        For j = 1 To vInstances
    
            For scan = iNumberOfAppointment - iNumberOfServers To iNumberOfAppointment - 1
                For k = scan + 1 To iNumberOfAppointment
                    If Tij(scan, j) > Tij(k, j) Then
                        dummy = Tij(scan, j)
                        Tij(scan, j) = Tij(k, j)
                        Tij(k, j) = dummy
                    End If
                    
                Next k
            Next scan
            
            Fi(iNumberOfAppointment + 1, j) = Tij(iNumberOfAppointment, j)
            Fis(j) = Fi(iNumberOfAppointment + 1, j)
        
        Next j
    
        Ai(iNumberOfAppointment + 1) = Application.WorksheetFunction.Average(Fis)
     
        For i = 1 To iNumberOfAppointment
            aAi(i) = aAi(i) + Ai(i) / dReplications  '(dStartOfDay * 60 + Ai(i)) / dReplications
        Next i
        
    Next o
            
    For i = 1 To iNumberOfAppointment - 1
            
        Cells(1 + i, 1).Value = i
        Cells(1 + i, 2).Value = aAi(i) / (24 * 60) 'appoitnemnt time per each appointment
        Cells(1 + i, 3).Value = aAi(i + 1) - aAi(i) ' slot
        Cells(1 + i, 4).Value = Di(i)
        Cells(1 + i, 5).Value = 1 - Si(i)
        Cells(1 + i, 6).Value = Ti(i)
        Cells(1 + i, 7).Value = sDistribution
    Next i
    
    Cells(1 + iNumberOfAppointment, 1).Value = iNumberOfAppointment
    Cells(1 + iNumberOfAppointment, 2).Value = aAi(iNumberOfAppointment) / (24 * 60)
    Cells(1 + iNumberOfAppointment, 3).Value = ""
    Cells(1 + iNumberOfAppointment, 4).Value = Di(iNumberOfAppointment)
    Cells(1 + iNumberOfAppointment, 5).Value = 1 - Si(iNumberOfAppointment)
    Cells(1 + iNumberOfAppointment, 6).Value = Ti(iNumberOfAppointment)
    Cells(1 + iNumberOfAppointment, 7).Value = sDistribution
'
'        i = i + 1
'        Cells(o + 1, i).Value = Mu 'writing the expected duration
    
'        i = i + 1
'        Cells(o + 1, i).Value = S 'writing the duration SD
    
'        i = i + 1
'        Cells(o + 1, i).Value = QoS 'writing the quality of service
    
 '       i = i + 1

        'calculating and writing the slot duration
 '       For j = 1 To 9
                
 '           Cells(o + 1, i).Value = Ai(j + 1) - Ai(j)
 '           i = i + 1
        
 '       Next j
        
        'Cells(o + 1, i).Value = "דר' יריב מרמור"

    'Next o
    Charts("Results").HasTitle = False
    Charts("Results").HasTitle = True
    
    If sDistribution = "Empirical Distribution" Then _
        Charts("Results").ChartTitle.Text = "Finding " & iNumberOfAppointment & " appointments based on:  " & Chr(10) & sDistribution & ", from file ''" & sFileName & " and " & iNumberOfServers & " servers" & Chr(10) & " (no-shows: " & dNoShow & "; QOS[P(Wait less than " & dWaitingTime & ")]=" & dQoS & ")" ''"
    If sDistribution = "Normal Distribution" Then _
                Charts("Results").ChartTitle.Text = "Finding " & iNumberOfAppointment & " appointments based on:  " & Chr(10) & sDistribution & ", with " & ChrW(956) & "=" & d_mu & ", " & ChrW(963) & "=" & d_sd & " and " & iNumberOfServers & " servers" & Chr(10) & " (no-shows: " & dNoShow & "; QOS[P(Wait less than " & dWaitingTime & ")]=" & dQoS & ")" ''"
    If sDistribution = "Uniform Distribution" Then _
                Charts("Results").ChartTitle.Text = "Finding " & iNumberOfAppointment & " appointments based on:  " & Chr(10) & sDistribution & ", with a=" & d_a & ", b=" & d_b & " and " & iNumberOfServers & " servers" & Chr(10) & " (no-shows: " & dNoShow & "; QOS[P(Wait less than " & dWaitingTime & ")]=" & dQoS & ")" ''"

    
   Charts("Results").Visible = True
   Charts("Results").Axes(xlValue, xlPrimary).MinimumScale = dStartOfDay / 24
   Charts("Results").Axes(xlValue, xlSecondary).MinimumScale = Ai(2) - Ai(1)
   Charts("Results").Activate
   'Charts("Results").Axes(xlValue).MinimumScale = [PriYMin]
    
    
End Sub

Sub fBubbleSort()
  Dim strTemp As String
  Dim i As Long
  Dim j As Long
  Dim lngMin As Long
  Dim lngMax As Long
  lngMin = LBound(dataArray)
  lngMax = UBound(dataArray)
  For i = lngMin To lngMax - 1
    For j = i + 1 To lngMax
      If dataArray(i) > dataArray(j) Then
        strTemp = dataArray(i)
        dataArray(i) = dataArray(j)
        dataArray(j) = strTemp
      End If
    Next j
  Next i
  
 ' lngMin = LBound(dataArray)
 ' lngMax = UBound(dataArray)
 ' For i = lngMin To lngMax - 1
 '   For j = i + 1 To lngMax
 '     If dataArray(i) > dataArray(j) Then
 '       strTemp = dataArray(i)
 '       dataArray(i) = dataArray(j)
 '       dataArray(j) = strTemp
 '     End If
 '   Next j
 ' Next i
End Sub
