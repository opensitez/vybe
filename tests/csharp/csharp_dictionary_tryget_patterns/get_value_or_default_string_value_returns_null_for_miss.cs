// vybe-test: csharp/csharp_dictionary_tryget_patterns/get_value_or_default_string_value_returns_null_for_miss
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, string>(); __Check((map.GetValueOrDefault(0) == null).ToString(), "True");
