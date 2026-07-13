use super::helpers::run_vb;

#[test]
fn string_interpolation_formatting() {
    let out = run_vb(
        r#"
Imports System.Globalization

Module M
    Sub Main()
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture
        
        Dim price As Decimal = 12.345D
        Dim pct As Double = 0.75
        
        ' Interpolation with formatting
        Console.WriteLine($"Price: {price:F2}")
        Console.WriteLine($"Percent: {pct:P0}")
        
        ' Interpolation with alignment
        Console.WriteLine($"[{price,10:F1}]")
        Console.WriteLine($"[{price,-10:F1}]")
    End Sub
End Module
"#,
    );
    assert_eq!(
        out,
        vec![
            "Price: 12.35",
            "Percent: 75 %",
            "[      12.3]",
            "[12.3      ]"
        ]
    );
}
