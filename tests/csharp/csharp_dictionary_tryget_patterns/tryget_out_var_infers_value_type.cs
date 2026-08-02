// vybe-test: csharp/csharp_dictionary_tryget_patterns/tryget_out_var_infers_value_type
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_tryget_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<string, int> { ["n"] = 7 }; if (map.TryGetValue("n", out var val)) __Check((val).ToString(), "7");
