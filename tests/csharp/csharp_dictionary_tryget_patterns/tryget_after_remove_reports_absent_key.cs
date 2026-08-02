// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_after_remove_reports_absent_key
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["gone"] = 5 }; map.Remove("gone"); __Check((map.TryGetValue("gone", out int v)).ToString(), "False");
