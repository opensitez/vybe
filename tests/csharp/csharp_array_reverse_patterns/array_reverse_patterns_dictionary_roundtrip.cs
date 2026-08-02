// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_reverse_patterns
var map = new System.Collections.Generic.Dictionary<int, int>(); map[27] = 28; __Check((map.ContainsKey(27) && map[27] == 28).ToString(), "True");
