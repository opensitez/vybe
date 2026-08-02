// vybe-test: csharp/csharp_jagged_array_patterns/jagged_array_patterns_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_jagged_array_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// jagged_array_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[28] = 29; __Check((map.ContainsKey(28) && map[28] == 29).ToString(), "True");
