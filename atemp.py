
Sub Button1_Click()



Dim counterLeft, counterRight, counterReport, LeftAccumulator, RightAccumulator As Integer
 

Dim MaxLimit As Integer

MaxLimit = Cells(2, 11).Value
counterLeft = 2
counterRight = 2
counterReport = 2

AccumulateLeft = True
AccumulateRight = True
LeftContentsIsOver = False
RightContentsIsOver = False

LeftAccumulator = 0
RightAccumulator = 0
    
'MsgBox (StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) <> 0)

'MsgBox (Cells(counterLeft, 2).Value)

'MsgBox ("-" + Cells(4, 5).Value + "-")

'MsgBox (StrComp(Cells(4, 5).Value, "", vbTextCompare = 0))

Range(Cells(2, 8), Cells(30000, 9)).Clear


While ( _
       (StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) <> 0) Or _
       (StrComp(Cells(counterRight, 5).Value, "", vbTextCompare) <> 0)) _
        And _
       (counterLeft < MaxLimit) And _
       (counterRight < MaxLimit)
   


        If AccumulateLeft Then
            LeftAccumulator = Cells(counterLeft, 3).Value
            While (Cells(counterLeft, 2).Value = Cells(counterLeft + 1, 2).Value) And _
                 (StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) <> 0)
            
                LeftAccumulator = LeftAccumulator + Cells(counterLeft + 1, 3).Value
                counterLeft = counterLeft + 1
            Wend
            AccumulateLeft = False
            
                        
        End If
        
        If AccumulateRight Then
            RightAccumulator = Cells(counterRight, 6).Value
            While ((Cells(counterRight, 5).Value = Cells(counterRight + 1, 5).Value) And _
                (StrComp(Cells(counterRight, 5).Value, "", vbTextCompare) <> 0))
                                
                RightAccumulator = RightAccumulator + Cells(counterRight + 1, 6).Value
                counterRight = counterRight + 1
            Wend
            AccumulateRight = False
            
             
           
            
            
        End If
               
        '''Compare
        If Cells(counterLeft, 2).Value = Cells(counterRight, 5).Value Then
            Cells(counterReport, 8).Value = Cells(counterLeft, 2).Value
            Cells(counterReport, 9).Value = RightAccumulator - LeftAccumulator
            
            counterReport = counterReport + 1
            
            
            counterLeft = counterLeft + 1
            counterRight = counterRight + 1
            
            
            If StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) = 0 Then
                    LeftContentsIsOver = True
            Else
                    AccumulateLeft = True
            End If
            If StrComp(Cells(counterRight, 5).Value, "", vbTextCompare) = 0 Then
                        RightContentsIsOver = True
            Else
                        AccumulateRight = True
            End If
            
        Else
            If ((Cells(counterLeft, 2).Value < Cells(counterRight, 5).Value) And _
                  StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) <> 0) _
                          Or _
                RightContentsIsOver Then
                Cells(counterReport, 8).Value = Cells(counterLeft, 2).Value
                Cells(counterReport, 9).Value = -1 * LeftAccumulator
                counterReport = counterReport + 1
                
                
                counterLeft = counterLeft + 1
                
                If StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) = 0 Then
                    LeftContentsIsOver = True
                Else
                    AccumulateLeft = True
                End If
            
        
            Else
                If (Cells(counterLeft, 2).Value > Cells(counterRight, 5).Value And _
                             StrComp(Cells(counterLeft, 2).Value, "", vbTextCompare) <> 0) Or _
                    LeftContentsIsOver Then
                 
                 Cells(counterReport, 8).Value = Cells(counterRight, 5).Value
                    Cells(counterReport, 9).Value = RightAccumulator
                    counterReport = counterReport + 1
                
                    
                    counterRight = counterRight + 1
                
'                    MsgBox (StrComp(Cells(counterRight, 5).Value, "", vbTextCompare = 0))
                    If StrComp(Cells(counterRight, 5).Value, "", vbTextCompare) = 0 Then
                        RightContentsIsOver = True
                    Else
                        AccumulateRight = True
                    End If
        
                End If
            End If
            
        End If
    
    
  Wend
    
    



If (counterLeft >= MaxLimit) Or (counterRight >= MaxLimit) Then
    MsgBox ("Error: No 'X' was provided at end of the list. Try again")
    
Else
    MsgBox ("Completed")
End If




End Sub



