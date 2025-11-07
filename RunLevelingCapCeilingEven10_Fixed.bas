'
' ============================================================================
' Module: RunLevelingCapCeilingEven10_Fixed
' Purpose:
'   Automates the distribution of troop unit counts in the `tblStacking`
'   Excel table so that:
'     1. Every troop keeps at least one unit when it participates in any cap.
'     2. The total resource cost for Leadership, Dominance, and Authority
'        stays within the per-stat limits defined on the worksheet.
'     3. Troops are leveled by total health so that higher-priority groups
'        (Specialists > Guardsmen > Monsters > Mercs) absorb as much health
'        as possible before lower-priority groups get additional units.
'     4. Extremely strong Mercs that would otherwise break the ordering are
'        left at a single “token” unit without forcing the Dominance cap to
'        overflow.
' Audience:
'   Written for workbook users who understand the game rules but are new to
'   VBA programming. Comments describe every significant step, variable, and
'   formula in plain English.
' ============================================================================
Option Explicit

' ============================================
' User-defined type to store troop information
' Each instance captures the row position in the table together with
' the per-unit health and per-unit cost for whichever resource governs it.
' ============================================
Private Type tTroop
    dHealth As Double   ' Health per unit of troop
    dCost As Double     ' Cost per unit for the governing stat
    iRow As Long        ' Row index within the table
End Type

' ============================================
' Helper: Safe string conversion
' Returns the supplied value as a trimmed string,
' or an empty string when the value is Null or an error.
' Parameters:
'   vValue - any worksheet value that might contain Null or an error.
' ============================================
Private Function NzS(vValue As Variant) As String
    If IsError(vValue) Then Exit Function      ' Guard against worksheet errors (e.g., #DIV/0!)
    If IsNull(vValue) Then Exit Function       ' Treat Null as empty
    NzS = Trim$(CStr(vValue))                  ' Convert to text and remove surrounding spaces
End Function

' ============================================
' Function: Find maximum target health H*
' Purpose:
'   Simulate “leveling” a set of troops to share a common target health (H*)
'   while respecting the remaining budget for the governing resource.
' Method:
'   Uses a binary search between 0 and a high guess to find the largest H*
'   that can be afforded when each troop either keeps its baseline units or
'   purchases extra units to reach H*.
' Parameters:
'   arrTroops()  - Array of troops belonging to the group we are leveling.
'   arrBase()    - Baseline unit counts (typically 1 or 0) for every worksheet row.
'   iCount       - Number of troops present in arrTroops().
'   dCap         - Remaining resource that can still be spent on this group.
' Returns:
'   The feasible target health H* expressed as “total health per troop”.
'   This value is later converted into an integer unit count per troop.
' Notes:
'   All mathematics stay in Double precision to avoid rounding errors during
'   the binary search, then shift to Long when determining unit counts.
' ============================================
Private Function FindLeadershipH(arrTroops() As tTroop, arrBase() As Long, iCount As Long, dCap As Double) As Double
    Dim dLow As Double          ' Lower bound for the binary search interval
    Dim dHigh As Double         ' Upper bound for the binary search interval
    Dim dMid As Double          ' Mid-point tested during the search
    Dim dBestRatio As Double    ' Highest health-per-cost ratio across troops
    Dim iIndex As Long          ' Loop counter for binary search iterations
    Dim dUsed As Double         ' Total resource cost required for a trial H*

    If iCount = 0 Or dCap <= 0 Then Exit Function    ' Nothing to do without troops or budget

    dLow = 0                    ' Start with the lowest possible target health
    dBestRatio = 0              ' Initialise the ratio used to estimate dHigh

    ' Step 1: locate the troop with the most efficient health-per-cost ratio.
    '         This ratio sets an optimistic upper bound for H*.
    For iIndex = 1 To iCount
        If arrTroops(iIndex).dCost > 0 And arrTroops(iIndex).dHealth > 0 Then
            dBestRatio = WorksheetFunction.Max(dBestRatio, arrTroops(iIndex).dHealth / arrTroops(iIndex).dCost)
        End If
    Next iIndex

    If dBestRatio <= 0 Then Exit Function            ' Exit early if every troop has zero cost or health

    dHigh = dCap * dBestRatio * 1.2                  ' Slightly inflated guess to ensure high is feasible

    ' Step 2: binary search. Forty iterations are plenty to converge for Double precision.
    For iIndex = 1 To 40
        dMid = 0.5 * (dLow + dHigh)                  ' Test the midpoint between low and high
        dUsed = 0                                    ' Track total cost to reach dMid

        Dim iLoop As Long
        For iLoop = 1 To iCount                      ' Evaluate each troop against the candidate H*
            If arrTroops(iLoop).dHealth > 0 Then
                Dim lNeeded As Long
                lNeeded = CLng(Fix(dMid / arrTroops(iLoop).dHealth))  ' Units required to hit dMid health
                If lNeeded > arrBase(arrTroops(iLoop).iRow) Then
                    ' Only additional units above the guaranteed baseline consume the remaining cap.
                    dUsed = dUsed + (lNeeded - arrBase(arrTroops(iLoop).iRow)) * arrTroops(iLoop).dCost
                End If
            End If
        Next iLoop

        ' Adjust the binary search anchors based on whether the candidate fits the budget.
        If dUsed <= dCap Then
            dLow = dMid            ' Candidate fits: move the lower bound upward to search for larger H*
        Else
            dHigh = dMid           ' Candidate is too expensive: shrink the upper bound
        End If
    Next iIndex

    FindLeadershipH = dLow        ' After the loop, dLow holds the highest affordable H*
End Function

' ============================================
' Procedure: AllocateLeadership
' Purpose:
'   Convert the target health H* returned by FindLeadershipH into integer
'   unit counts for every troop in the group. Baseline units are preserved
'   so that nobody drops below their guaranteed minimum.
' Parameters:
'   arrTroops() - Troops to update.
'   arrBase()   - Baseline units by worksheet row.
'   iCount      - Number of troops contained in arrTroops().
'   arrUnits()  - Output array; will hold the updated unit totals.
'   dHstar      - The target health H* that the group should attempt to reach.
' ============================================
Private Sub AllocateLeadership(arrTroops() As tTroop, arrBase() As Long, iCount As Long, arrUnits() As Long, dHstar As Double)
    Dim iIndex As Long             ' Loop counter through each troop in the group

    If iCount = 0 Or dHstar <= 0 Then Exit Sub    ' Nothing to allocate

    For iIndex = 1 To iCount
        If arrTroops(iIndex).dHealth > 0 Then
            Dim lNeeded As Long
            lNeeded = CLng(Fix(dHstar / arrTroops(iIndex).dHealth))      ' Integer units to reach H*
            If lNeeded < arrBase(arrTroops(iIndex).iRow) Then            ' Never drop below the baseline
                arrUnits(arrTroops(iIndex).iRow) = arrBase(arrTroops(iIndex).iRow)
            Else
                arrUnits(arrTroops(iIndex).iRow) = lNeeded
            End If
        Else
            arrUnits(arrTroops(iIndex).iRow) = arrBase(arrTroops(iIndex).iRow)  ' Zero-health troops keep baseline
        End If
    Next iIndex
End Sub

' ============================================
' Helper: compute resource usage for a group
' Adds up the total cost of the units currently assigned to a troop array.
' Parameters:
'   arrTroops() - Troops governed by the resource we want to measure.
'   iCount      - How many troops are in the array.
'   arrUnits()  - Current unit counts for every worksheet row.
' Returns:
'   The overall cost spent on this troop array.
' ============================================
Private Function ComputeUsedCost(arrTroops() As tTroop, iCount As Long, arrUnits() As Long) As Double
    Dim iIndex As Long            ' Loop counter across troops
    Dim dTotal As Double          ' Running cost total

    For iIndex = 1 To iCount
        dTotal = dTotal + arrUnits(arrTroops(iIndex).iRow) * arrTroops(iIndex).dCost
    Next iIndex

    ComputeUsedCost = dTotal
End Function

' ============================================
' Helper: ceiling division for positive numbers
' Purpose:
'   Given a target value and a step size (typically troop health),
'   determine how many whole steps are needed to meet or exceed the target.
' Parameters:
'   dValue - Target total health we wish to reach.
'   dStep  - Health contributed by a single unit.
' Returns:
'   The minimum integer number of units required.
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
    Dim iHighIdx As Long, iLowIdx As Long   ' Indexes of the weakest “high” troop and strongest “low” troop
    Dim dMinHigh As Double, dMaxLow As Double  ' Health totals for the selected high/low troops
    Dim lUnits As Long                     ' Current unit count for the troop being examined
    Dim dTotal As Double                   ' Current total health for the troop being examined
    Dim lNeeded As Long                    ' Units the high troop needs to overtake the low troop
    Dim lAdd As Long                       ' Additional units granted to the high troop
    Dim dUsage As Double                   ' Current cost already spent by the high troop group
    Dim dRemain As Double                  ' Spare resource headroom still available
    Dim lAffordable As Long                ' How many more units the high group can afford
    Dim lRowHigh As Long                   ' Worksheet row of the high troop being boosted
    Dim lGuard As Long                     ' Safety counter to avoid infinite loops
    Dim arrIgnored() As Boolean            ' Flags to skip low troops that are “stuck” at their baseline

    If iHigh = 0 Or iLow = 0 Or dCap <= 0 Then Exit Sub
    If iLow > 0 Then
        ReDim arrIgnored(1 To iLow)
    End If

    For lGuard = 1 To 10000
        iHighIdx = 0: iLowIdx = 0
        dMinHigh = 0: dMaxLow = 0

        ' Locate the weakest high-priority troop (the smallest total health).
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

        ' Locate the strongest low-priority troop (the largest total health),
        ' ignoring troops that are fixed at their baseline.
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

        If iHighIdx = 0 Or iLowIdx = 0 Then Exit For     ' No viable comparison left
        If dMinHigh > dMaxLow + 0.0001 Then Exit For     ' Priority already satisfied
        If arrHigh(iHighIdx).dCost <= 0 Then             ' High troop cannot consume resource; ignore this low troop
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

        dUsage = ComputeUsedCost(arrHigh, iHigh, arrUnits)   ' Current spend on the high-priority group
        dRemain = dCap - dUsage                              ' Remaining budget for that resource
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

        arrUnits(lRowHigh) = arrUnits(lRowHigh) + lAdd       ' Add affordable units to push the high troop ahead
ContinueEnsure:
    Next lGuard
End Sub

' ============================================
' Helper: enforce health priority (high vs low group)
' Purpose:
'   Ensure that every troop in a higher-priority group has total health that
'   meets or exceeds every troop in the lower-priority group, without breaking
'   the governing resource cap. When a lower-priority troop is “stuck” at its
'   minimum unit count, the routine marks it as ignored so the macro keeps the
'   troop at one unit but does not continue to chase an impossible ordering.
' Parameters:
'   arrHigh()          - Higher-priority troop records (e.g., Specialists vs Guardsmen).
'   iHigh              - How many troops exist in arrHigh().
'   arrLow(), iLow     - Lower-priority troop records and their count.
'   arrUnits()         - Current unit counts (shared across the whole routine).
'   arrBase()          - Baseline units per worksheet row.
'   lStep              - Adjustment tick size (1 pre-rounding, 10 after rounding).
'   arrResource(), iResource, dResourceCap
'                      - Reference group used to measure resource spend. For example,
'                        when comparing Monsters vs Mercs, the Monster array is needed
'                        to calculate Dominance usage.
'   bUseResource       - When True, the routine checks that additional units fit within
'                        the resource cap before allowing them.
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
' Workflow summary:
'   1. Read the worksheet limits and classify each row into the four troop tiers.
'   2. Guarantee a minimum of one unit for every troop that consumes any resource.
'   3. Reserve that baseline cost, then binary-search the remaining cap for each tier
'      (Leadership → Dominance → Authority) to produce a shared health target H*.
'   4. Enforce the priority rules so higher tiers stay ahead in total health.
'   5. Optionally round unit counts to the nearest 10 while preserving ordering.
'   6. Write the results back to the table and (optionally) report resource usage.
' ============================================
Public Sub RunLevelingCapCeilingEven10_Fixed()
    Dim ws As Worksheet                               ' Active worksheet containing caps and the troop table
    Dim dLcap As Double, dDcap As Double, dAcap As Double ' Worksheet cap inputs for Leadership, Dominance, Authority
    Dim loTable As ListObject                         ' Table object representing tblStacking
    Dim iRows As Long                                 ' Number of troop rows in the table
    Dim arrUnits() As Long                            ' Final unit counts to write back

    Dim arrSpec() As tTroop, arrGuard() As tTroop     ' Leadership sub-groups by priority
    Dim arrMonster() As tTroop, arrMerc() As tTroop   ' Dominance and Authority groups
    Dim arrLeadership() As tTroop                     ' Combined Leadership troops (Specialists + Guardsmen)
    Dim iSpec As Long, iGuard As Long, iMonster As Long, iMerc As Long
    Dim iLeadership As Long                           ' Counts for each troop array
    Dim arrBase() As Long                             ' Baseline “at least one unit” values per row

    Dim dHLeadership As Double, dHMonster As Double, dHMerc As Double  ' Target health H* values per tier
    Dim dBaseLeadership As Double, dBaseDominance As Double, dBaseAuthority As Double ' Cost of the baseline units
    Dim dLcapAvail As Double, dDcapAvail As Double, dAcapAvail As Double              ' Caps remaining after baseline

    Dim dUsedL_pre As Double, dUsedD_pre As Double, dUsedA_pre As Double              ' Cost before rounding
    Dim dUsedL_post As Double, dUsedD_post As Double, dUsedA_post As Double           ' Cost after rounding
    Dim dLostL As Double, dLostD As Double, dLostA As Double                          ' Percentage loss due to rounding

    Dim iIndex As Long                           ' General-purpose loop counter

    ' ----------------------------------------------------------------------
    ' Step 1: gather worksheet references and cap values.
    ' ----------------------------------------------------------------------
    Set ws = ThisWorkbook.ActiveSheet
    Set loTable = ws.ListObjects("tblStacking")

    dLcap = ws.Range("Leadership_Cap").Value
    dDcap = ws.Range("Dominance_Cap").Value
    dAcap = ws.Range("Authority_Cap").Value

    iRows = loTable.ListRows.Count
    ReDim arrUnits(1 To iRows)
    ReDim arrBase(1 To iRows)

    iSpec = 0: iGuard = 0: iMonster = 0: iMerc = 0

    Dim iTypeCol As Long                             ' Column index for tblStacking[Type]
    Dim iMercCol As Long                             ' Column index for tblStacking[Merc]
    Dim sType As String                              ' Text value of the troop type
    Dim bIsMerc As Boolean                           ' True when the row is marked as a Merc

    iTypeCol = loTable.ListColumns("Type").Index
    iMercCol = loTable.ListColumns("Merc").Index

    ' ----------------------------------------------------------------------
    ' Step 2: read each row, determine its troop category, and store the
    '         baseline unit guarantee (1 if it uses any resource, otherwise 0).
    ' ----------------------------------------------------------------------
    For iIndex = 1 To iRows
        Dim rRow As ListRow
        Set rRow = loTable.ListRows(iIndex)

        Dim dHealth As Double                         ' Total health per unit for this troop
        Dim dLc As Double, dDc As Double, dAc As Double   ' Per-unit costs for each stat

        ' NzD is a workbook helper (not shown here) that converts empty cells to 0.0.
        dHealth = NzD(rRow.Range(1, loTable.ListColumns("Health").Index).Value)
        dLc = NzD(rRow.Range(1, loTable.ListColumns("Leadership").Index).Value)
        dDc = NzD(rRow.Range(1, loTable.ListColumns("Dominance").Index).Value)
        dAc = NzD(rRow.Range(1, loTable.ListColumns("Authority").Index).Value)

        If dHealth > 0 And (dLc > 0 Or dDc > 0 Or dAc > 0) Then
            arrBase(iIndex) = 1          ' Guarantee at least one unit whenever the troop consumes any resource
        Else
            arrBase(iIndex) = 0          ' Otherwise the troop starts at zero and may remain unused
        End If

        sType = UCase$(NzS(rRow.Range(1, iTypeCol).Value))
        bIsMerc = (UCase$(NzS(rRow.Range(1, iMercCol).Value)) = "Y")

        If bIsMerc Then
            If dAc > 0 Then
                ' Mercenaries consume Authority and belong to the lowest priority tier.
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
                        ' Specialists spend Leadership and sit at the top of the priority list.
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
                        ' Guardsmen also spend Leadership but rank below Specialists.
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
                        ' Monsters consume Dominance and form the third tier.
                        iMonster = iMonster + 1
                        ReDim Preserve arrMonster(1 To iMonster)
                        arrMonster(iMonster).dHealth = dHealth
                        arrMonster(iMonster).dCost = dDc
                        arrMonster(iMonster).iRow = iIndex
                    End If
                Case Else
                    If dLc > 0 Then
                        ' Fallback: treat unexpected labels that still spend Leadership as Guardsmen.
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
            ' Record the cost of the guaranteed baseline units so we know how much cap is already spoken for.
            dBaseLeadership = dBaseLeadership + arrBase(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Leadership").Index).Value)
            dBaseDominance = dBaseDominance + arrBase(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Dominance").Index).Value)
            dBaseAuthority = dBaseAuthority + arrBase(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Authority").Index).Value)
        End If
    Next iIndex

    ' Remaining budget after honouring the baseline allocations.
    dLcapAvail = dLcap - dBaseLeadership
    If dLcapAvail < 0 Then dLcapAvail = 0
    dDcapAvail = dDcap - dBaseDominance
    If dDcapAvail < 0 Then dDcapAvail = 0
    dAcapAvail = dAcap - dBaseAuthority
    If dAcapAvail < 0 Then dAcapAvail = 0

    ' ----------------------------------------------------------------------
    ' Step 3: run the leveling solver for each tier to find target health H*.
    ' ----------------------------------------------------------------------
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

    ' ----------------------------------------------------------------------
    ' Step 4: enforce the health ordering between tiers before rounding.
    ' ----------------------------------------------------------------------
    EnforcePriority arrSpec, iSpec, arrGuard, iGuard, arrUnits, arrBase, 1, arrLeadership, iLeadership, dLcap, True
    EnforcePriority arrGuard, iGuard, arrMonster, iMonster, arrUnits, arrBase, 1, arrLeadership, iLeadership, dLcap, True
    EnforcePriority arrMonster, iMonster, arrMerc, iMerc, arrUnits, arrBase, 1, arrMonster, iMonster, dDcap, True
    EnsureHigherBeatsLower arrMonster, iMonster, arrMerc, iMerc, arrUnits, arrBase, dDcap

    For iIndex = 1 To iRows
        dUsedL_pre = dUsedL_pre + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Leadership").Index).Value)
        dUsedD_pre = dUsedD_pre + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Dominance").Index).Value)
        dUsedA_pre = dUsedA_pre + arrUnits(iIndex) * NzD(loTable.ListRows(iIndex).Range(1, loTable.ListColumns("Authority").Index).Value)
    Next iIndex

    Const bApplyRounding As Boolean = True        ' Toggle: set to False to skip the “nearest 10” rounding pass

    If bApplyRounding Then
        ' ------------------------------------------------------------------
        ' Step 5: round unit counts to multiples of 10 (minimum 10) where possible.
        ' ------------------------------------------------------------------
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

        ' Re-check ordering after rounding to ensure it still holds.
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
    ' ----------------------------------------------------------------------
    ' Step 6: push the computed unit counts back into the worksheet.
    ' ----------------------------------------------------------------------
    For iIndex = 1 To iRows
        loTable.ListRows(iIndex).Range(1, iUnitCol).Value = arrUnits(iIndex)
    Next iIndex

    ' Calculate how much resource was effectively “lost” because of rounding.
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
