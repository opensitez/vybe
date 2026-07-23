use super::helpers::run_vb;

#[test]
fn cultureinfo_invariant_values_are_stable() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim invariant As CultureInfo = CultureInfo.InvariantCulture

        Console.WriteLine(invariant.Name)
        Console.WriteLine(invariant.IsNeutralCulture)
        Console.WriteLine(invariant.NumberFormat.NumberDecimalSeparator)
        Console.WriteLine(invariant.NumberFormat.CurrencySymbol)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["en-US", "False", ".", "$"]);
}

#[test]
fn cultureinfo_parse_and_format_with_invariant_decimal() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim us As CultureInfo = CultureInfo.GetCultureInfo("en-US")
        Dim parsed As Double = Double.Parse("1.25", us)
        Console.WriteLine(parsed)
        Console.WriteLine(parsed.ToString("F2", CultureInfo.InvariantCulture))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1.5", "1.25"]);
}

#[test]
fn cultureinfo_currency_format_is_predictable() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim value As Decimal = 12D
        Dim text As String = value.ToString("C", CultureInfo.GetCultureInfo("en-US"))
        Console.WriteLine(text.StartsWith("$"))
        Console.WriteLine(text.Contains("12"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn cultureinfo_datetime_parse_and_fields_are_accessible() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

    Module M
    Sub Main()
        Dim dt As DateTime = DateTime.Parse("07/21/2026", CultureInfo.GetCultureInfo("en-US"))
        Console.WriteLine(dt.Year)
        Console.WriteLine(dt.Month)
        Console.WriteLine(dt.Day)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["2026", "7", "21"]);
}

#[test]
fn cultureinfo_neutral_and_specific_relationship() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim specific As CultureInfo = CultureInfo.GetCultureInfo("en-US")
        Dim parentName As String = specific.Parent.Name

        Console.WriteLine(specific.IsNeutralCulture)
        Console.WriteLine(parentName = "en")
        Console.WriteLine(specific.Equals(CultureInfo.GetCultureInfo("en-US")))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["False", "True", "True"]);
}

#[test]
fn cultureinfo_text_transforms_are_stable() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim ti As TextInfo = CultureInfo.GetCultureInfo("en-US").TextInfo

        Console.WriteLine(ti.ToTitleCase("hello world"))
        Console.WriteLine(ti.ToUpper("vb"))
        Console.WriteLine(ti.ToLower("VB"))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["Hello World", "VB", "vb"]);
}

#[test]
fn cultureinfo_enumerates_neutral_cultures() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim cultures() As CultureInfo = CultureInfo.GetCultures(CultureTypes.NeutralCultures)
        Console.WriteLine(cultures.Length > 0)
        Console.WriteLine(cultures.Length > 5)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn cultureinfo_read_only_property_is_true_for_invariant() {
    let out = run_vb(
        r#"
Imports System
Imports System.Globalization

Module M
    Sub Main()
        Dim culture As CultureInfo = CultureInfo.InvariantCulture
        Dim clone As CultureInfo = CType(culture.Clone(), CultureInfo)

        Console.WriteLine(culture.IsReadOnly)
        Console.WriteLine(Not clone.IsReadOnly)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}
