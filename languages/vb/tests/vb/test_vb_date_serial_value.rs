use super::helpers::run_vb;

#[test]
fn date_serial_value() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' DateSerial creates a date from year, month, day
        Dim d = DateSerial(2022, 10, 15)
        Console.WriteLine(d.Year)
        Console.WriteLine(d.Month)
        
        ' DateValue parses a string
        Dim dv = DateValue("2023-05-20")
        Console.WriteLine(dv.Day)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2022", "10", "20"]);
}
