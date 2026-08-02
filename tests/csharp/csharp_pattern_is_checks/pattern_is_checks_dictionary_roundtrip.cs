// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[41] = 42; __Check((map.ContainsKey(41) && map[41] == 42).ToString(), "True");
