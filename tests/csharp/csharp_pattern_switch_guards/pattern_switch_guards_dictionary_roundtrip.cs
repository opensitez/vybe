// vybe-test: csharp/csharp_pattern_switch_guards/pattern_switch_guards_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_switch_guards.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_switch_guards
var map = new System.Collections.Generic.Dictionary<int, int>(); map[42] = 43; __Check((map.ContainsKey(42) && map[42] == 43).ToString(), "True");
