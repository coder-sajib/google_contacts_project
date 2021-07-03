' See README.MD

Name_Phone_and_Email_Cleanup_and_Extraction()
Text = ""

' object for current active sheet, works irrestpective of sheetname
Set CurrentSheet = ActiveSheet

rangevalue = 2500 ' get the range for searching  mobile no

Rows2 = CurrentSheet.UsedRange.Rows.Count
' code  for searching the mobile no fields'
Set colvalue = Range("A1:DP" & rangevalue).Find(What:="Mobile", LookAt:=xlWhole, MatchCase:=False, SearchFormat:=False)

Phone1colvalue = colvalue.Column + 1
Phone2colvalue = Phone1colvalue + 2
Phone3colvalue = Phone1colvalue + 4
Phone4colvalue = Phone1colvalue + 6
Phone5colvalue = Phone1colvalue + 8
Phone6colvalue = Phone1colvalue + 10
Phone7colvalue = Phone1colvalue + 12
Phone8colvalue = Phone1colvalue + 14
Phone9colvalue = Phone1colvalue + 16
Phone10colvalue = Phone1colvalue + 18
Phone11colvalue = Phone1colvalue + 20
Phone12colvalue = Phone1colvalue + 22

Call WrtitiOutput("----------------------------------------------------------------Updated Contacts on " & Now() & "----------------------------------------------------------------")

Dim Texobj As String


For i = 2 To Rows2 + 10
    
    
    If Rows(i).Hidden = False Then        ' checking filtered values
    
                                                                            For j = Phone1colvalue - 1 To Phone12colvalue
                                                                              CurrentSheet.Cells(i, j).Select
                                                                              
                                                                                                PhoneValue = CurrentSheet.Cells(i, j).Value
                                                                                                Firstnam = CurrentSheet.Cells(i, 1).Value
                                                                                               
                                                                                                
                                                                                                 If InStr(1, PhoneValue, "1") > 0 Or InStr(1, PhoneValue, "2") > 0 Or InStr(1, PhoneValue, "3") > 0 Or InStr(1, PhoneValue, "4") > 0 Or InStr(1, PhoneValue, "5") > 0 Or InStr(1, PhoneValue, "6") > 0 Or InStr(1, PhoneValue, "7") > 0 Or InStr(1, PhoneValue, "8") > 0 Or InStr(1, PhoneValue, "9") > 0 Or InStr(1, PhoneValue, "0") > 0 Then

                                                                                                    
                                                                                                    TextHasNoValue = "Yes"
                                                                                                    Else
                                                                                                    
                                                                                                     TextHasNoValue = "No"
                                                                                                    End If
 

                                                                                             If Firstnam <> Empty And PhoneValue <> Empty And TextHasNoValue = "Yes" And Len(PhoneValue) > 9 Then
                                                                                                    
                                                                                                                              
                                                                                                                                
                                                                                                                        
                                                                                                                                                        
                                                                                                                                                        If InStr(1, PhoneValue, ":") > 0 Then
                                                                                                                                                            
                                                                                                                                                            PhoneValue1 = Replace(PhoneValue, ":::", ":")
                                                                                                                                                            PhoneValue1 = Split(PhoneValue1, ":")
                                                                                                                                                            CurrentSheet.Cells(i, j).Value = PhoneValue1(0)
                                                                                                                                                            CurrentSheet.Cells(i, j + 2).Value = PhoneValue1(1)
                                                                                                                                                            Texobj = CurrentSheet.Cells(i, 1).Value & "," & PhoneValue1(0)
                                                                                                                                                             Call WrtitiOutput(Texobj)
                                                                                                                                                            Texobj = CurrentSheet.Cells(i, 1).Value & "," & PhoneValue1(1)
                                                                                                                                                            Call WrtitiOutput(Texobj)
                                                                                                                                                          
                                                                                                                                                        Else
                                                                                                                                                            'code for replacing special characters
                                                                                                                                                            
                                                                                                                                                            PhoneValue1 = Replace(PhoneValue, "(", "")
                                                                                                                                                            PhoneValue1 = Replace(PhoneValue1, ")", "")
                                                                                                                                                            PhoneValue2 = Replace(PhoneValue1, "-", "")
                                                                                                                                                            PhoneValue2 = Replace(PhoneValue1, "-", "")
                                                                                                                                                            FilterNo = Replace(PhoneValue2, " ", "")
                                                                                                                                                            
                                                                                                                                                            LenofFilterno = Len(FilterNo)        ' gets the length
                                                                                                                                                            
                                                                                                                                                                        If LenofFilterno > 10 And Left(FilterNo, 1) = "+" Then        ' check if the number  with  1xxxxxxxxxx format
                                                                                                                                                                                    
                                                                                                                                                                                    If InStr(1, PhoneValue2, " ") > 0 Then
                                                                                                                                                                                        
                                                                                                                                                                                        FilterNo1 = Split(PhoneValue2, " ")
                                                                                                                                                                                        FilterNo12 = FilterNo1(0)
                                                                                                                                                                                        FilterNo13 = FilterNo1(1)
                                                                                                                                                                                        
                                                                                                                                                                                        FilterNo14 = Right(PhoneValue2, Len(PhoneValue2) - (Len(FilterNo12) + 11))        ' geting text value in phone no
                                                                                                                                                                                        FilterNo5 = Replace(Replace(FilterNo, FilterNo14, ""), " ", "") ' Removing text from  phone no
                                                                
                                                                                                                                                                                        CurrentSheet.Cells(i, j).Value = FilterNo5        'Entering the updated phone no  in excel
                                                                                                                                                                                        CurrentSheet.Cells(i, j + 1).Value = FilterNo14        ' updating the text in next column
                                                                                                                                                                                        Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo5
                                                                                                                                                                                          Call WrtitiOutput(Texobj)
                                                                                                                                                            
                                                                                                                                                                                    Else
                                                                                                                                                                                        
                                                                                                                                                                                        CurrentSheet.Cells(i, j).Value = FilterNo
                                                                                                                                                                                        Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo
                                                                                                                                                                                         Call WrtitiOutput(Texobj)
                                                                                                                                                                                        
                                                                                                                                                                                    End If
                                                                                                                                                                       
                                                                                                                                                    
                                                                                                                                    ElseIf LenofFilterno > 13 And Left(FilterNo, 1) <> "+" Then
                                                                                                                                        
                                                                                                                                                If InStr(1, FilterNo, "+") < 3 Then
                                                                                                                                                 
                                                                                                                                    
                                                                                                                                          
                                                                                                                                                         If InStr(1, PhoneValue2, " ") > 0 And InStr(1, PhoneValue2, ",") = 0 Then
                                                                                                                                                          
                                                                                                                                                                  FilterNo1 = Split(PhoneValue2, " ")
                                                                                                                                                                  FilterNo12 = FilterNo1(0)
                                                                                                                                                                  
                                                                                                                                                                  'geting text value in phone no
                                                                                                                                                                  
                                                                                                                                                                  FilterNo14 = Right(PhoneValue2, Len(PhoneValue2) - (Len(FilterNo12) + 8))
                                                                                                                                                                  FilterNo5 = Replace(FilterNo, FilterNo14, "")        ' Removing text from  phone no
                                                                                                                                                                 ' FilterNo5 = Replace(PhoneValue2, FilterNo14, "")
                                                                                                                                                                   FilterNo5 = Replace(Replace(PhoneValue2, FilterNo14, ""), " ", "") ' Removing text from  phone no
                                                                    
                                                                                                                                                                  CurrentSheet.Cells(i, j).Value = FilterNo5        'Entering the updated phone no  in excel
                                                                                                                                                                  CurrentSheet.Cells(i, j + 1).Value = FilterNo14        ' updating the text in next column
                                                                                                                                                                    Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo5
                                                                                                                                                                 Call WrtitiOutput(Texobj)
                                                                                                                                                          ElseIf InStr(1, PhoneValue2, ",") = 0 Then
                                                                                                                                                          
                                                                                                                                                                      FilterNo1 = Split(PhoneValue2, " ")
                                                                                                                                                                      FilterNo12 = FilterNo1(0)
                                                                                                                                                                      
                                                                                                                                                                      FilterNo14 = Right(PhoneValue2, Len(PhoneValue2) - (Len(FilterNo12) - 4))        'geting text value in phone no FilterNo5 = Replace(FilterNo, FilterNo14, "")  ' Removing text from  phone no
                                                                                                                                                                      FilterNo5 = Replace(PhoneValue2, FilterNo14, "")
                                                                                                                                                                      FilterNo5 = Replace(Replace(PhoneValue2, FilterNo14, ""), " ", "") ' Removing text from  phone no
                                                                                                                                                                      
                                                                                                                                                                      CurrentSheet.Cells(i, j).Value = FilterNo5        'Entering the updated phone no  in excel
                                                                                                                                                                      CurrentSheet.Cells(i, j + 1).Value = FilterNo14        ' updating the text in next column
                                                                                                                                                                      'updating the text in next column
                                                                                                                                                                       
                                                                                                                                                                       Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo5
                                                                                                                                                                       Call WrtitiOutput(Texobj)
                                                                                                                                                           End If
                                                                                                                                                         
                                                                                                                                                        Else
                                                                                                                                                        
                                                                                                                                                                  
                                                                                                                                                                                       FilterNo1 = Split(FilterNo, "+")
                                                                                                                                                                                         FilterNo12 = FilterNo1(0)
                                                                                                                                                                                       FilterNo13 = FilterNo1(1)
                                                                                                                                                                                        
                                                                                                                                                                                         
                                                                                                                                                                                      
                                                                                                                                                                                         CurrentSheet.Cells(i, j).Activate
                                                                                                                                                                                  MsgBox ("'+" & FilterNo13 & FilterNo12)
                                                                                                                                                                                  CurrentSheet.Cells(i, j).Value = "'+" & FilterNo13 & FilterNo12    ' updating the text in next column
                                                                                                                                                                                        Texobj = CurrentSheet.Cells(i, 1).Value & "," & "+" & FilterNo13 & " " & FilterNo12
                                                                                                                                                                                          Call WrtitiOutput(Texobj)
                                                                                                                                                         
                                                                                                                                                               
                                                                                                                                                               End If
                                                                                                                                                               
                                                                                                                                                        
                                                                                                                                      ElseIf LenofFilterno = 12 And Left(FilterNo, 1) <> "+" Then
                                                                                                                                        
                                                                                                                                                       If InStr(1, PhoneValue2, " ") > 0 And InStr(1, PhoneValue2, ",") = 0 Then
                                                                                                                                                        
                                                                                                                                                                FilterNo1 = Split(PhoneValue2, " ")
                                                                                                                                                                FilterNo12 = FilterNo1(0)
                                                                                                                                                                
                                                                                                                                                                'geting text value in phone no
                                                                                                                                                                
                                                                                                                                                                FilterNo14 = Right(PhoneValue2, Len(PhoneValue2) - (Len(FilterNo12) + 8))
                                                                                                                                                                FilterNo5 = Replace(FilterNo, FilterNo14, "")        ' Removing text from  phone no
                                                                                                                                                               ' FilterNo5 = Replace(PhoneValue2, FilterNo14, "")
                                                                                                                                                                 FilterNo5 = Replace(Replace(PhoneValue2, FilterNo14, ""), " ", "") ' Removing text from  phone no
                                                                  
                                                                                                                                                                CurrentSheet.Cells(i, j).Value = FilterNo5        'Entering the updated phone no  in excel
                                                                                                                                                                CurrentSheet.Cells(i, j + 1).Value = FilterNo14        ' updating the text in next column
                                                                                                                                                                  Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo5
                                                                                                                                                               Call WrtitiOutput(Texobj)
                                                                                                                                                        
                                                                                                                                                        ElseIf InStr(1, PhoneValue2, ",") = 0 Then
                                                                                                                                                        
                                                                                                                                                        
                                                                                                                                                                    FilterNo1 = Split(PhoneValue2, " ")
                                                                                                                                                                    FilterNo12 = FilterNo1(0)
                                                                                                                                                                    
                                                                                                                                                                    FilterNo14 = Right(PhoneValue2, Len(PhoneValue2) - (Len(FilterNo12) - 4))        'geting text value in phone no FilterNo5 = Replace(FilterNo, FilterNo14, "")  ' Removing text from  phone no
                                                                                                                                                                    FilterNo5 = Replace(PhoneValue2, FilterNo14, "")
                                                                                                                                                                    FilterNo5 = Replace(Replace(PhoneValue2, FilterNo14, ""), " ", "") ' Removing text from  phone no
                                                                                                                                                                    
                                                                                                                                                                    CurrentSheet.Cells(i, j).Value = FilterNo5        'Entering the updated phone no  in excel
                                                                                                                                                                    CurrentSheet.Cells(i, j + 1).Value = FilterNo14        ' updating the text in next column
                                                                                                                                                                    'updating the text in next column
                                                                                                                                                                     
                                                                                                                                                                     Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo5
                                                                                                                                                                     Call WrtitiOutput(Texobj)
                                                                                                                                                         End If
                                                                                                                                                      
                                                                                                                                         
                                                                                                                                        
                                                                                                                                      ElseIf LenofFilterno > 8 And LenofFilterno < 11 Then
                                                                                                                                        
                                                                                                                                        CurrentSheet.Cells(i, j).Value = "1" & FilterNo
                                                                                                                                        
                                                                                                                                         Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo
                                                                                                                                        Call WrtitiOutput(Texobj)
                                                                                                                                        
                                                                                                                                        Else
                                                                                                                                         Texobj = CurrentSheet.Cells(i, 1).Value & "," & FilterNo
                                                                                                                                        Call WrtitiOutput(Texobj)
                                                                                                                            
                                                                                                                            
                                                                                                                                       End If
                                                                                                                                    
                                                                                                                                    
                                                                                                       
                                                                                                        
                                                                                                        
                                                                                 
                                                                                                
                                                                                                
                                                     
                                                                                                                                                        
                                                                                            
                                                                            
                                                                        
                                                                        
                                                                            

          
                                                                                                                                        End If
                                                                                                                                                                          
                                                                                         
 
                                                                                                                                                            
                                                                                                                                                              
                                                                                                                                                          
                                                                                             End If


                Next

          End If

Next
MsgBox ("Format Completed")
 
 
 End Sub
 
 Function WrtitiOutput(Texobj As String)
                                                        If Texobj <> "" Then
                                                        
                                                        
                                                            strFile_Path = "C:\Users\" & Environ("UserName") & "\Desktop\Output2.txt" 'Change as per your test folder and exiting file path to append it.
                                                            Open strFile_Path For Append As #1
                                                            Print #1, Texobj
                                                            Close #1
                                                            
                                                            End If
                                                            
 End Function
