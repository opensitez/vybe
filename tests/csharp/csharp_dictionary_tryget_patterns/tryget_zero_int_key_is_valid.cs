// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_zero_int_key_is_valid
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, int> { [0] = 0 }; __Check((map.TryGetValue(0, out int v)).ToString(), "True"); __Check((v).ToString(), "0");
