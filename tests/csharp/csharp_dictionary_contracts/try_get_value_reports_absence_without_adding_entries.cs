// vybe-test: csharp/csharp_dictionary_contracts/try_get_value_reports_absence_without_adding_entries
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var map = new Dictionary<string, int> { ["a"] = 1 };
bool found = map.TryGetValue("missing", out var value);
__Check((found ? "Y" : "N").ToString(), "N");
__Check((map.Count).ToString(), "1");
