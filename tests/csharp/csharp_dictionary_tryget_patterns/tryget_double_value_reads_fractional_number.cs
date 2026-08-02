// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_double_value_reads_fractional_number
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, double> { ["pi"] = 3.14 }; map.TryGetValue("pi", out double d); __Check((d).ToString(), "3.14");
