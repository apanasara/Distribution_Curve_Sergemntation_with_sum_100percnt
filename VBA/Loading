Option Explicit

Public Enum Loading_Type
    Front = -1
    Back = 1
End Enum

Private Function Loading(Step As Integer, Span As Integer, Degree As Integer) As Double
    
    Span = Span + 1 'Loading upto end point
    
    '----https://stackoverflow.com/questions/53074786/front-loaded-and-back-loaded-normal-distribution-column-chart-and-s-curves-in
    'BackLoading = Sin(ratio + (Sin(ratio) / Degree))
   
    Dim delta As Integer, i As Integer
    Dim ratio As Double, LSum As Double, Li As Double
        
    Dim Pi As Double: Pi = 22 / 7
    
        
    For i = 0 To Span
    
        Select Case i
            Case 0: Li = 0 ' Start Correction
            Case Span: Li = 0 ' End Correction
            Case Else
                delta = Span - i
                ratio = Pi * delta / Span
                Li = Sin(ratio + (Sin(ratio) / Degree))
        End Select

        LSum = LSum + Li
        If i = Step Then: Loading = Li
    Next
    
    'scaling to 100%
    Loading = Loading / LSum
    
End Function


Public Function BackLoading(Step As Integer, Span As Integer, Degree As Integer) As Double
    BackLoading = Loading(Step, Span, Degree * Back)
End Function


Public Function FrontLoading(Step As Integer, Span As Integer, Degree As Integer) As Double
    FrontLoading = Loading(Step, Span, Degree * Front)
End Function

Public Function NormalLoading(Step As Integer, Span As Integer, Degree As Integer) As Double
    NormalLoading = Application.WorksheetFunction.Norm_Dist(Step, Span / 2, Span / Degree, False)
End Function
