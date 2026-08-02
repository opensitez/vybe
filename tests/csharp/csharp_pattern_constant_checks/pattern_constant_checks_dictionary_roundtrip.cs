// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[40] = 41; __Check((map.ContainsKey(40) && map[40] == 41).ToString(), "True");
