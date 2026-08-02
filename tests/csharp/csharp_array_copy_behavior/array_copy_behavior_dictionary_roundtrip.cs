// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
var map = new System.Collections.Generic.Dictionary<int, int>(); map[26] = 27; __Check((map.ContainsKey(26) && map[26] == 27).ToString(), "True");
