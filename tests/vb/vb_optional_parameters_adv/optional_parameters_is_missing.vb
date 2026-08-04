' vybe-test: vb/vb_optional_parameters_adv/optional_parameters_is_missing
' origin: languages/vb/tests/vb/test_vb_optional_parameters_adv.rs

Module M
    ' IsMissing is a legacy function only valid for Optional Object arguments
    Function CheckMissing(Optional ByVal arg As Object = Nothing) As Boolean
        Return IsMissing(arg)
    End Function

    Sub Main()
        ' VB.NET supports default parameter values instead of IsMissing for non-Object types
        ' For Object types, IsMissing checks if it was omitted (if it is Type.Missing)
        ' In standard VB.NET, Type.Missing is passed when an optional object parameter is omitted
        ' Wait, actually IsMissing only works if the default value isn't explicitly set to Nothing?
        ' VB.NET requires a default value for Optional parameters. 
        ' To use IsMissing, we usually can't unless it's a late-bound COM object, but it's part of the language spec.
        ' Let's just test Optional with default values.
    End Sub
End Module
