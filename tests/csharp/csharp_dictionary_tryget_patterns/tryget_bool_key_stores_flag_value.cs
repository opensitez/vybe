// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_bool_key_stores_flag_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<bool, string> { [true] = "yes" }; map.TryGetValue(true, out string s); __Check((s).ToString(), "yes");
