// vybe-test: csharp/csharp_type_conversions/is_operator_reports_true_for_assignable_interface
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; object item = new List<int>(); __Check((item is IEnumerable<int>).ToString(), "True");
