// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[115] = 116; __Check((map.ContainsKey(115) && map[115] == 116).ToString(), "True");
