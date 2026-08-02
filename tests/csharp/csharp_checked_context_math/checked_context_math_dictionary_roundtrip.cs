// vybe-test: csharp/csharp_checked_context_math/checked_context_math_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_checked_context_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// checked_context_math
var map = new System.Collections.Generic.Dictionary<int, int>(); map[12] = 13; __Check((map.ContainsKey(12) && map[12] == 13).ToString(), "True");
