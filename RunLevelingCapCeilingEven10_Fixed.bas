Option Explicit

' ============================================
' User-defined type to store troop information
' ============================================
Private Type tTroop
    dHealth As Double   ' Health per unit of troop
    dCost As Double     ' Cost per unit for the governing stat
    iRow As Long        ' Row index within the table
End Type

' ============================================
' Helper: Safe string conversion
' ============================================
Private Function NzS(vValue As Variant) As String
    If IsError(vValue) Then Exit Function
    If IsNull(vValue) Then Exit Function
    NzS = Trim$(CStr(vValue))
End Function

' ============================================
' Function: Find maximum target health H*
' using binary search for a troop group
' ============================================
Private Function FindLeadershipH(arrTroops() As tTroop, arrBase() As Long, iCount As Long, dCap As Double) As Double
    Dim dLow As Double
    Dim dHigh As Double
    Dim dMid As Double
    Dim dBestRatio As Double
    Dim iIndex As Long
    Dim dUsed As Double

    If iCount = 0 Or dCap <= 0 Then Exit Function

    dLow = 0
    dBestRatio = 0

    For iIndex = 1 To iCount
        If arrTroops(iIndex).dCost > 0 And arrTroops(iIndex).dHealth > 0 Then
            dBestRatio = WorksheetFunction.Max(dBestRatio, arrTroops(iIndex).dHealth / arrTroops(iIndex).dCost)
        End If
    Next iIndex

    If dBestRatio <= 0 Then Exit Function

    dHigh = dCap * dBestRatio * 1.2

    For iIndex = 1 To 40
        dMid = 0.5 * (dLow + dHigh)
        dUsed = 0

        Dim iLoop As Long
        For iLoop = 1 To iCount
            If arrTroops(iLoop).dHealth > 0 Then
                Dim lNeeded As Long
                lNeeded = CLng(Fix(dMid / arrTroops(iLoop).dHealth))
                If lNeeded > arrBase(arrTroops(iLoop).iRow) Then
                    dUsed = dUsed + (lNeeded - arrBase(arrTroops(iLoop).iRow)) * arrTroops(iLoop).dCost
                End If
            End If
        Next iLoop

        If dUsed <= dCap Then
            dLow = dMid
        Else
            dHigh = dMid
        End If
    Next iIndex

    FindLeadershipH = dLow
End Function

' ============================================
' Allocate units for a troop group
' ============================================
Private Sub AllocateLeadership(arrTroops() As tTroop, arrBase() As Long, iCount As Long, arrUnits() As Long, dHstar As Double)
    Dim iIndex As Long

    If iCount = 0 Or dHstar <= 0 Then Exit Sub

    For iIndex = 1 To iCount
        If arrTroops(iIndex).dHealth > 0 Then
            Dim lNeeded As Long
            lNeeded = CLng(Fix(dHstar / arrTroops(iIndex).dHealth))
            If lNeeded < arrBase(arrTroops(iIndex).iRow) Then
                arrUnits(arrTroops(iIndex).iRow) = arrBase(arrTroops(iIndex).iRow)
            Else
                arrUnits(arrTroops(iIndex).iRow) = lNeeded
            End If
        Else
            arrUnits(arrTroops(iIndex).iRow) = arrBase(arrTroops(iIndex).iRow)
        End If
    Next iIndex
End Sub

' ============================================
' Helper: compute resource usage for a group
' ============================================
Private Function ComputeUsedCost(arrTroops() As tTroop, iCount As Long, arrUnits() As Long) As Double
    Dim iIndex As Long
    Dim dTotal As Double

    For iIndex = 1 To iCount
        dTotal = dTotal + arrUnits(arrTroops(iIndex).iRow) * arrTroops(iIndex).dCost
    Next iIndex

    ComputeUsedCost = dTotal
End Function

' ============================================
' Helper: ceiling division for positive numbers
' ============================================
Private Function CeilToLong(ByVal dValue As Double, ByVal dStep As Double) As Long
    If dStep <= 0 Then Exit Function
    CeilToLong = CLng(-Int(-(dValue / dStep)))
    If CeilToLong < 0 Then CeilToLong = 0
End Function

' ============================================
' Helper: increase higher-priority troops until they exceed lower tier health
' ============================================
Private Sub EnsureHigherBeatsLower(arrHigh() As tTroop, ByVal iHigh As Long, _
                                   arrLow() As tTroop, ByVal iLow As Long, _
                                   arrUnits() As Long, arrBase() As Long, _
                                   ByVal dCap As Double)

    Dim iIndex As Long
    Dim iHighIdx As Long, iLowIdx As Long
    Dim dMinHigh As Double, dMaxLow As Double
    Dim lUnits As Long
    Dim dTotal As Double
    Dim lNeeded As Long
    Dim lAdd As Long
    Dim dUsage As Double
    Dim dRemain As Double
    Dim lAffordable As Long
    Dim lRowHigh As Long
    Dim lGuard As Long
    Dim arrIgnored() As Boolean

    If iHigh = 0 Or iLow = 0 Or dCap <= 0 Then Exit Sub
    If iLow > 0 Then
        ReDim arrIgnored(1 To iLow)
    End If

    For lGuard = 1 To 10000
        iHighIdx = 0: iLowIdx = 0
        dMinHigh = 0: dMaxLow = 0

        For iIndex = 1 To iHigh
            lUnits = arrUnits(arrHigh(iIndex).iRow)
            If lUnits > 0 And arrHigh(iIndex).dHealth > 0 Then
                dTotal = lUnits * arrHigh(iIndex).dHealth
                If iHighIdx = 0 Or dTotal < dMinHigh Then
                    dMinHigh = dTotal
                    iHighIdx = iIndex
                End If
            End If
        Next iIndex

        For iIndex = 1 To iLow
            If Not arrIgnored(iIndex) Then
                lUnits = arrUnits(arrLow(iIndex).iRow)
                If lUnits > 0 And arrLow(iIndex).dHealth > 0 Then
                    dTotal = lUnits * arrLow(iIndex).dHealth
                    If dTotal > dMaxLow Then
                        dMaxLow = dTotal
                        iLowIdx = iIndex
                    End If
                End If
            End If
        Next iIndex

        If iHighIdx = 0 Or iLowIdx = 0 Then Exit For
        If dMinHigh > dMaxLow + 0.0001 Then Exit For
        If arrHigh(iHighIdx).dCost <= 0 Then
            arrIgnored(iLowIdx) = True
            GoTo ContinueEnsure
        End If

        lRowHigh = arrHigh(iHighIdx).iRow
        lUnits = arrUnits(lRowHigh)
        lNeeded = CeilToLong(dMaxLow + 0.0001, arrHigh(iHighIdx).dHealth)
        If lNeeded < arrBase(lRowHigh) Then lNeeded = arrBase(lRowHigh)
        If lNeeded <= lUnits Then lNeeded = lUnits + 1
        lAdd = lNeeded - lUnits
        If lAdd <= 0 Then lAdd = 1

        dUsage = ComputeUsedCost(arrHigh, iHigh, arrUnits)
        dRemain = dCap - dUsage
        If dRemain <= 0 Then
            arrIgnored(iLowIdx) = True
            GoTo ContinueEnsure
        End If

        lAffordable = CLng(Fix(dRemain / arrHigh(iHighIdx).dCost))
        If lAffordable <= 0 Then
            arrIgnored(iLowIdx) = True
            GoTo ContinueEnsure
        End If
        If lAdd > lAffordable Then lAdd = lAffordable
        If lAdd <= 0 Then
            arrIgnored(iLowIdx) = True
            GoTo ContinueEnsure
        End If

        arrUnits(lRowHigh) = arrUnits(lRowHigh) + lAdd
ContinueEnsure:
    Next lGuard
End Sub

' ============================================
' Helper: enforce health priority (high vs low group)
' ============================================
Private Sub EnforcePriority(arrHigh() As tTroop, ByVal iHigh As Long, _
                            arrLow() As tTroop, ByVal iLow As Long, _
                            arrUnits() As Long, arrBase() As Long, _
                            ByVal lStep As Long, _
                            arrResource() As tTroop, ByVal iResourceCount As Long, _
                            ByVal dResourceCap As Double, ByVal bUseResource As Boolean)

    Dim iIndex As Long
    Dim dTotal As Double
    Dim lUnits As Long
    Dim dMinHigh As Double
    Dim dMaxLow As Double
    Dim lTargetRow As Long
    Dim lHighRow As Long
    Dim bHasHigh As Boolean
    Dim lDecrement As Long
    Dim lIncrement As Long
    Dim dCostPerUnit As Double
    Dim dCurrentUsage As Double
    Dim lMaxAdd As Long
    Dim lGuard As Long
    Dim arrIgnored() As Boolean
    Dim lLowIdx As Long
    Dim lHighIdx As Long

    If iHigh = 0 Or iLow = 0 Then Exit Sub
    If lStep < 1 Then lStep = 1
    If Not bUseResource Or iResourceCount <= 0 Or dResourceCap < 0 Then
        bUseResource = False
    End If
    If iLow > 0 Then
        ReDim arrIgnored(1 To iLow)
    End If

    For lGuard = 1 To 10000
        bHasHigh = False
        dMinHigh = 0
        lHighRow = 0
        lHighIdx = 0

        For iIndex = 1 To iHigh
            lUnits = arrUnits(arrHigh(iIndex).iRow)
            If lUnits > 0 And arrHigh(iIndex).dHealth > 0 Then
                dTotal = lUnits * arrHigh(iIndex).dHealth
                If Not bHasHigh Or dTotal < dMinHigh Then
                    dMinHigh = dTotal
                    lHighRow = arrHigh(iIndex).iRow
                    lHighIdx = iIndex
                End If
                bHasHigh = True
            End If
        Next iIndex

        If Not bHasHigh Then Exit Sub

        dMaxLow = 0
        lTargetRow = 0
        lLowIdx = 0

        For iIndex = 1 To iLow
            If Not arrIgnored(iIndex) Then
                lUnits = arrUnits(arrLow(iIndex).iRow)
                If lUnits > 0 And arrLow(iIndex).dHealth > 0 Then
                    dTotal = lUnits * arrLow(iIndex).dHealth
                    If dTotal > dMaxLow Then
                        dMaxLow = dTotal
                        lTargetRow = arrLow(iIndex).iRow
                        lLowIdx = iIndex
                    End If
                End If
            End If
        Next iIndex

        If lTargetRow = 0 Then Exit Sub
        If dMaxLow < dMinHigh - 0.0000001 Then Exit For

        If arrUnits(lTargetRow) <= arrBase(lTargetRow) Then
            Dim bBoosted As Boolean
            bBoosted = False
            If lHighRow <> 0 Then
                If lStep > 1 Then
                    lIncrement = lStep
                Else
                    lIncrement = 1
                End If
                If bUseResource Then
                    dCostPerUnit = arrHigh(lHighIdx).dCost
                    If dCostPerUnit > 0 Then
                        dCurrentUsage = ComputeUsedCost(arrResource, iResourceCount, arrUnits)
                        lMaxAdd = CLng(Fix((dResourceCap - dCurrentUsage) / dCostPerUnit))
                        If lMaxAdd < 0 Then lMaxAdd = 0
                        If lStep > 1 Then
                            If lIncrement > lMaxAdd Then lIncrement = (lMaxAdd \ lStep) * lStep
                        Else
                            If lIncrement > lMaxAdd Then lIncrement = lMaxAdd
                        End If
                    End If
                End If
                If lIncrement > 0 Then
                    arrUnits(lHighRow) = arrUnits(lHighRow) + lIncrement
                    bBoosted = True
                End If
            End If
            If Not bBoosted Then
                If lLowIdx > 0 Then arrIgnored(lLowIdx) = True
            End If
            GoTo ContinueLoop
        End If

        If lStep > 1 And arrUnits(lTargetRow) - lStep >= arrBase(lTargetRow) Then
            lDecrement = lStep
        Else
            lDecrement = 1
        End If

        If arrUnits(lTargetRow) - lDecrement < arrBase(lTargetRow) Then
            Dim bRaised As Boolean
            bRaised = False
            If lHighRow <> 0 Then
                If lStep > 1 Then
                    lIncrement = lStep
                Else
                    lIncrement = 1
                End If
                If bUseResource Then
                    dCostPerUnit = arrHigh(lHighIdx).dCost
                    If dCostPerUnit > 0 Then
                        dCurrentUsage = ComputeUsedCost(arrResource, iResourceCount, arrUnits)
                        lMaxAdd = CLng(Fix((dResourceCap - dCurrentUsage) / dCostPerUnit))
                        If lMaxAdd < 0 Then lMaxAdd = 0
                        If lStep > 1 Then
                            If lIncrement > lMaxAdd Then lIncrement = (lMaxAdd \ lStep) * lStep
                        Else
                            If lIncrement > lMaxAdd Then lIncrement = lMaxAdd
                        End If
                    End If
                End If
                If lIncrement > 0 Then
                    arrUnits(lHighRow) = arrUnits(lHighRow) + lIncrement
                    bRaised = True
                End If
            End If
            If Not bRaised Then
                If lLowIdx > 0 Then arrIgnored(lLowIdx) = True
            End If
            GoTo ContinueLoop
        End If
        If lDecrement <= 0 Then Exit For

        arrUnits(lTargetRow) = arrUnits(lTargetRow) - lDecrement
ContinueLoop:
    Next lGuard
End Sub

' ============================================
' Main routine: Distribute units across groups
' ============================================
Public Sub RunLevelingCapCeilingEven10_Fixed()
    Dim ws As Worksheet
    Dim dLcap As Double, dDcap As Double, dAcap As Double
    Dim loTable As ListObject
    Dim iRows As Long
    Dim arrUnits() As Long

    Dim arrSpec() As tTroop, arrGuard() As tTroop
    Dim arrMonster() As tTroop, arrMerc() As tTroop
    Dim arrLeadership() As tTroop
    Dim iSpec As Long, iGuard As Long, iMonster As Long, iMerc As Long
    Dim iLeadership As Long
    Dim arrBase() As Long

    Dim dHLeadership As Double, dHMonster As Double, dHMerc As Double
    Dim dBaseLeadership As Double, dBaseDominance As Double, dBaseAuthority As Double
    Dim dLcapAvail As Double, dDcapAvail As Double, dAcapAvail As Double

    Dim dUsedL_pre As Double, dUsedD_pre As Double, dUsedA_pre As Double
    Dim dUsedL_post As Double, dUsedD_post As Double, dUsedA_post As Double
    Dim dLostL As Double, dLostD As Double, dLostA As Double

    Dim iIndex As Long

    Set ws = ThisWorkbook.ActiveSheet
    Set loTable = ws.ListObjects("tblStacking")

    dLcap = ws.Range("Leadership_Cap").Value
    dDcap = ws.Range("Dominance_Cap").Value
    dAcap = ws.Range("Authority_Cap").Value

    iRows = loTable.ListRows.Count
    ReDim arrUnits(1 To iRows)
    ReDim arrBase(1 To iRows)

    iSpec = 0: iGuard = 0: iMonster = 0: iMerc = 0

    Dim iTypeCol As Long
    Dim iMercCol As Long
    Dim sType As String
    Dim bIsMerc As Boolean

    iTypeCol = loTable.ListColumns("Type").Index
    iMercCol = loTable.ListColumns("Merc").Index

    For iIndex = 1 To iRows
        Dim rRow As ListRow
        Set rRow = loTable.ListRows(iIndex)

        Dim dHealth As Double
        Dim dLc As Double, dDc As Double, dAc As Double

        dHealth = NzD(rRow.Range(1, loTable.ListColumns("Health").Index).Value)
        dLc = NzD(rRow.Range(1, loTable.ListColumns("Leadership").Index).Value)
        dDc = NzD(rRow.Range(1, loTable.ListColumns("Dominance").Index).Value)
        dAc = NzD(rRow.Range(1, loTable.ListColumns("Authority").Index).Value)

        If dHealth > 0 And (dLc > 0 Or dDc > 0 Or dAc > 0) Then
            arrBase(iIndex) = 1
        Else
            arrBase(iIndex) = 0
        End If

        sType = UCase$(NzS(rRow.Range(1, iTypeCol).Value))
        bIsMerc = (UCase$(NzS(rRow.Range(1, iMercCol).Value)) = "Y")

        If bIsMerc Then
            If dAc > 0 Then
                iMerc = iMerc + 1
                ReDim Preserve arrMerc(1 To iMerc)
                arrMerc(iMerc).dHealth = dHealth
                arrMerc(iMerc).dCost = dAc
                arrMerc(iMerc).iRow = iIndex
            End If
        Else
            Select Case sType
                Case "SPECIALIST"
                    If dLc > 0 Then
                        iSpec = iSpec + 1
                        ReDim Preserve arrSpec(1 To iSpec)
                        arrSpec(iSpec).dHealth = dHealth
                        arrSpec(iSpec).dCost = dLc
                        arrSpec(iSpec).iRow = iIndex

                        iLeadership = iLeadership + 1
                        ReDim Preserve arrLeadership(1 To iLeadership)
                        arrLeadership(iLeadership).dHealth = dHealth
                        arrLeadership(iLeadership).dCost = dLc
                        arrLeadership(iLeadership).iRow = iIndex
                    End If
                Case "GUARDSMAN"
                    If dLc > 0 Then
                        iGuard = iGuard + 1
                        ReDim Preserve arrGuard(1 To iGuard)
                        arrGuard(iGuard).dHealth = dHealth
                        arrGuard(iGuard).dCost = dLc
                        arrGuard(iGuard).iRow = iIndex

                        iLeadership = iLeadership + 1
                        ReDim Preserve arrLeadership(1 To iLeadership)
                        arrLeadership(iLeadership).dHealth = dHealth
                        arrLeadership(iLeadership).dCost = dLc
                        arrLeadership(iLeadership).iRow = iIndex
                    End If
                Case "MONSTER"
                    If dDc > 0 Then
                        iMonster = iMonster + 1
                        ReDim Preserve arrMonster(1 To iMonster)
                        arrMonster(iMonster).dHealth = dHealth
                        arrMonster(iMonster).dCost = dDc
                        arrMonster(iMonster).iRow = iIndex
                    End If
                Case Else
                    If dLc > 0 Then
                        iGuard = iGuard + 1
                        ReDim Preserve arrGuard(1 To iGuard)
                        arrGuard(iGuard).dHealth = dHealth
                        arrGuard(iGuard).dCost = dLc
                        arrGuard(iGuard).iRow = iIndex

                        iLeadership = iLeadership + 1
                        ReDim Preserve arrLeadership(1 To iLeadership)
                        arrLeadership(iLeadership).dHealth = dHealth
                        arrLeadership(iLeadership).dCost = dLc
                        arrLeadership(iLeadership).iRow = iIndex
                    End If
            End Select

        End If
    Next iIndex

    For iIndex = 1 To iRows
        arrUnits(iIndex) = arrBase(iIndex)
        If arrBase(iIndex) > 0 Then
            dBaseLeadership = dBaseLeadership + arrBase(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Leadership").Index).Value)
            dBaseDominance = dBaseDominance + arrBase(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Dominance").Index).Value)
            dBaseAuthority = dBaseAuthority + arrBase(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Authority").Index).Value)
        End If
    Next iIndex

    dLcapAvail = dLcap - dBaseLeadership
    If dLcapAvail < 0 Then dLcapAvail = 0
    dDcapAvail = dDcap - dBaseDominance
    If dDcapAvail < 0 Then dDcapAvail = 0
    dAcapAvail = dAcap - dBaseAuthority
    If dAcapAvail < 0 Then dAcapAvail = 0

    If iLeadership > 0 Then
        dHLeadership = FindLeadershipH(arrLeadership, arrBase, iLeadership, dLcapAvail)
        AllocateLeadership arrLeadership, arrBase, iLeadership, arrUnits, dHLeadership
    End If

    If iMonster > 0 Then
        dHMonster = FindLeadershipH(arrMonster, arrBase, iMonster, dDcapAvail)
        AllocateLeadership arrMonster, arrBase, iMonster, arrUnits, dHMonster
    End If

    If iMerc > 0 And dAcap > 0 Then
        dHMerc = FindLeadershipH(arrMerc, arrBase, iMerc, dAcapAvail)
        AllocateLeadership arrMerc, arrBase, iMerc, arrUnits, dHMerc
    End If

    EnforcePriority arrSpec, iSpec, arrGuard, iGuard, arrUnits, arrBase, 1, arrLeadership, iLeadership, dLcap, True
    EnforcePriority arrGuard, iGuard, arrMonster, iMonster, arrUnits, arrBase, 1, arrLeadership, iLeadership, dLcap, True
    EnforcePriority arrMonster, iMonster, arrMerc, iMerc, arrUnits, arrBase, 1, arrMonster, iMonster, dDcap, True
    EnsureHigherBeatsLower arrMonster, iMonster, arrMerc, iMerc, arrUnits, arrBase, dDcap

    For iIndex = 1 To iRows
        dUsedL_pre = dUsedL_pre + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Leadership").Index).Value)
        dUsedD_pre = dUsedD_pre + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Dominance").Index).Value)
        dUsedA_pre = dUsedA_pre + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Authority").Index).Value)
    Next iIndex

    Const bApplyRounding As Boolean = True

    If bApplyRounding Then
        For iIndex = 1 To iRows
            If arrUnits(iIndex) <= 0 Then
                arrUnits(iIndex) = 1
            Else
                Dim iRounded As Long
                iRounded = (arrUnits(iIndex) \ 10) * 10
                If iRounded >= 10 Then
                    arrUnits(iIndex) = iRounded
                End If
            End If
        Next iIndex

        EnforcePriority arrSpec, iSpec, arrGuard, iGuard, arrUnits, arrBase, 10, arrLeadership, iLeadership, dLcap, True
        EnforcePriority arrGuard, iGuard, arrMonster, iMonster, arrUnits, arrBase, 10, arrLeadership, iLeadership, dLcap, True
        EnforcePriority arrMonster, iMonster, arrMerc, iMerc, arrUnits, arrBase, 10, arrMonster, iMonster, dDcap, True
        EnsureHigherBeatsLower arrMonster, iMonster, arrMerc, iMerc, arrUnits, arrBase, dDcap
    End If

    For iIndex = 1 To iRows
        dUsedL_post = dUsedL_post + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Leadership").Index).Value)
        dUsedD_post = dUsedD_post + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Dominance").Index).Value)
        dUsedA_post = dUsedA_post + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Authority").Index).Value)
    Next iIndex

    Dim iUnitCol As Long
    iUnitCol = loTable.ListColumns("Base Qty").Index
    For iIndex = 1 To iRows
        loTable.ListRows(iIndex).Range(1, iUnitCol).Value = arrUnits(iIndex)
    Next iIndex

    If dUsedL_pre > 0 Then dLostL = (dUsedL_pre - dUsedL_post) / dUsedL_pre * 100
    If dUsedD_pre > 0 Then dLostD = (dUsedD_pre - dUsedD_post) / dUsedD_pre * 100
    If dUsedA_pre > 0 Then dLostA = (dUsedA_pre - dUsedA_post) / dUsedA_pre * 100

    ' Optional: show summary
    ' MsgBox "Troop leveling complete." & vbCrLf & vbCrLf & _
    '        "Leadership used: " & Format(dUsedL_pre, "#,##0") & " -> " & Format(dUsedL_post, "#,##0") & _
    '        "   (Lost " & Format(dLostL, "0.00") & "%)" & vbCrLf & _
    '        "Dominance used:  " & Format(dUsedD_pre, "#,##0") & " -> " & Format(dUsedD_post, "#,##0") & _
    '        "   (Lost " & Format(dLostD, "0.00") & "%)" & vbCrLf & _
    '        "Authority used:  " & Format(dUsedA_pre, "#,##0") & " -> " & Format(dUsedA_post, "#,##0") & _
    '        "   (Lost " & Format(dLostA, "0.00") & "%)" & vbCrLf & vbCrLf & _
    '        "H* (Leadership/Dominance/Authority): " & Format(dHLeadership, "#,##0") & " / " & _
    '        Format(dHMonster, "#,##0") & " / " & Format(dHMerc, "#,##0"), vbInformation
End Sub
