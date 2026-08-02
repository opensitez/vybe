// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_bool_stored_value
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, bool> { ["flag"] = true }; map.TryGetValue("flag", out bool b); __Check((b).ToString(), "True");
