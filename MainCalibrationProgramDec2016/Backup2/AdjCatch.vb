﻿Module AdjCatchMod
    Sub adjcatch()
        '**********************************************************************
        'Subroutine to adjust recoveries in each fishery based upon the value
        ' of CatchFlag%:
        '    0 = No adjustment;
        '    1 = Adjust model catch to equal estimated catch;
        '    2 = Adjust model catch to equal estimated catch if model catch
        '        exceeds estimated catch.
        '    3 = use external adjustment
        '**********************************************************************

        'COMPUTE ADJUSTMENT FACTOR FOR EACH FISHERY
        ReDim RecAdjFactor(NumFish)
        ReDim FishExpCWTAll(NumStk, MaxAge, NumFish, NumSteps)

        Dim counter As Integer = 0
        Dim fishstring As String = " "
        Dim fishstringall As String = " "

        Dim CWTCatch As String
        CWTCatch = "C:\data\calibration\07Qbasic\CWTCatch"
        FileOpen(25, CWTCatch, OpenMode.Output)

        Dim Unadjusted As String
        Unadjusted = filepath & "\Unadjusted.txt"
        FileOpen(16, Unadjusted, OpenMode.Output)

        Print(16, "Fishery" & "," & "AnnualCatch" & "," & "TrueCatch" & "," & vbCrLf)




        For Fish = 1 To NumFish

            Print(16, Fish & "," & AnnualCatch(Fish) & "," & TrueCatch(Fish) & vbCrLf)

            If Firstpass = True Then
                CatchFlag(Fish) = 0
            End If

            If NoExpansions = True Then
                CatchFlag(Fish) = 0
            End If


            Select Case CatchFlag(Fish) 'located in cal file to the right of base period fishery catches or in BasePeriodCatch table of Calibration Support db

                ' ADJUST MODEL CATCH TO ESTIMATE CATCH


                Case 1
                    If AnnualCatch(Fish) > 0 Then
                        RecAdjFactor(Fish) = TrueCatch(Fish) / AnnualCatch(Fish) 'true catch = base period catch 
                    Else
                        counter = counter + 1
                        fishstring = CStr(Fish)
                        fishstringall = fishstringall & ", " & fishstring
                        'MsgBox("Base period catch exists yet recoveries are zero for fishery " & Fish)
                    End If
                    For TStep = 1 To NumSteps
                        For STk = MinStk To NumStk
                            For Age = 2 To MaxAge
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish)
                                End If
                            Next Age
                        Next STk
                    Next TStep

                    ' ADJUST MODEL CATCH IF GREATER THAN ESTIMATED CATCH

                Case 2
                    If AnnualCatch(Fish) > 0 Then
                        RecAdjFactor(Fish) = TrueCatch(Fish) / AnnualCatch(Fish)
                    Else
                        counter = counter + 1
                        fishstring = CStr(Fish)
                        fishstringall = fishstringall & ", " & fishstring
                        'MsgBox("Base period catch exists yet recoveries are zero for fishery " & Fish)
                    End If
                    If RecAdjFactor(Fish) < 1 Then
                        For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish)
                                    End If
                                Next Age
                            Next STk
                        Next TStep
                    Else
                        RecAdjFactor(Fish) = 99

                        For TStep = 1 To NumSteps
                            For STk = MinStk To NumStk
                                For Age = 2 To MaxAge
                                    If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                        FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                                    End If
                                Next Age
                            Next STk
                        Next TStep
                    End If

                    ' ADJUST MODEL CATCH TO ESTIMATE CATCH


                Case 3
                    If AnnualCatch(Fish) > 0 Then
                        RecAdjFactor(Fish) = TrueCatch(Fish) * ExternalModelStockProportion(Fish) / AnnualCatch(Fish) 'true catch = base period catch 
                    Else
                        counter = counter + 1
                        fishstring = CStr(Fish)
                        fishstringall = fishstringall & ", " & fishstring
                        'MsgBox("Base period catch exists yet recoveries are zero for fishery " & Fish)
                    End If
                    For TStep = 1 To NumSteps
                        For STk = MinStk To NumStk
                            For Age = 2 To MaxAge
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep) * RecAdjFactor(Fish)
                                End If
                            Next Age
                        Next STk
                    Next TStep


                    'NO ADJUSTMENT

                Case Else
                    For TStep = 1 To NumSteps
                        For STk = MinStk To NumStk
                            For Age = 2 To MaxAge
                                If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                    FishExpCWTAll(STk, Age, Fish, TStep) = CWTAll(STk, Age, Fish, TStep)
                                End If
                            Next Age
                        Next STk
                    Next TStep
                    RecAdjFactor(Fish) = 99
            End Select

            RecAdjFactor(Fish) = TrueCatch(Fish) / AnnualCatch(Fish)

        Next Fish
        FileClose(16)

        AdjustedCatch = True

        If counter > 0 Then
            MsgBox("Base period catch exists yet recoveries are zero for fishery " & fishstringall)
        End If
        If Firstpass <> True Then
            Print(25, "Stock,  Age, Fish, TStep, CWT" & vbCrLf)
            For STk = 1 To NumStk
                'Print(25, "ESCExpansionFactor " & "," & STk & "," & EscExpFact(STk) & vbCrLf)
                For Age = 2 To MaxAge
                    For Fish = 1 To NumFish
                        For TStep = 1 To NumSteps
                            If CWTAll(STk, Age, Fish, TStep) > 0 Then
                                Print(25, STk & "," & Age & "," & Fish & "," & TStep & "," & FishExpCWTAll(STk, Age, Fish, TStep) * EscExpFact(STk) & vbCrLf)
                            End If
                        Next
                    Next
                Next
            Next
        End If

    End Sub
End Module
