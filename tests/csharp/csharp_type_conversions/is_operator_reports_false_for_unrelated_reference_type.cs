// vybe-test: csharp/csharp_type_conversions/is_operator_reports_false_for_unrelated_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = "text"; __Check((item is System.DateTime).ToString(), "False");
