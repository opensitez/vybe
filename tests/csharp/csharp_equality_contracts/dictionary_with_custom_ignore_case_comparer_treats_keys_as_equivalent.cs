// vybe-test: csharp/csharp_equality_contracts/dictionary_with_custom_ignore_case_comparer_treats_keys_as_equivalent
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
var map = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
map["User"] = 1;
__Check((map.ContainsKey("user")).ToString(), "True");
__Check((map["USER"]).ToString(), "1");
