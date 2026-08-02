// vybe-test: csharp/csharp_partial_type_behavior/partial_type_behavior_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_partial_type_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// partial_type_behavior
var map = new System.Collections.Generic.Dictionary<int, int>(); map[70] = 71; __Check((map.ContainsKey(70) && map[70] == 71).ToString(), "True");
