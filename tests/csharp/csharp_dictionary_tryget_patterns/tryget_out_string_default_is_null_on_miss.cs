// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_out_string_default_is_null_on_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string>(); map.TryGetValue(1, out string s); __Check((s == null).ToString(), "True");
