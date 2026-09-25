Option Explicit

' modBlendEngine
' Fixed-split contribution logic for the blended factor portfolio.
' Weights are applied at each entry/contribution date. There is NO ongoing
' rebalancing after a contribution - each sleeve simply follows its own path.

' Lumpsum: growth factor of one rupee split across the two sleeves at entry.
Public Function BlendedGrowthFactor(ByVal momStart As Double, ByVal momEnd As Double, _
                                    ByVal valStart As Double, ByVal valEnd As Double, _
                                    ByVal wMom As Double, ByVal wVal As Double) As Double
    BlendedGrowthFactor = wMom * (momEnd / momStart) + wVal * (valEnd / valStart)
End Function

' SIP: units accumulated in each sleeve, valued at maturity prices.
Public Function BlendedSIPMaturity(ByVal unitsMom As Double, ByVal unitsVal As Double, _
                                   ByVal momEnd As Double, ByVal valEnd As Double) As Double
    BlendedSIPMaturity = unitsMom * momEnd + unitsVal * valEnd
End Function

' Split a single contribution into the Momentum and Value sleeves.
Public Sub ContributionSplit(ByVal amount As Double, ByVal wMom As Double, ByVal wVal As Double, _
                             ByRef amtMom As Double, ByRef amtVal As Double)
    amtMom = amount * wMom
    amtVal = amount * wVal
End Sub
