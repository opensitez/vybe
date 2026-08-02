// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[18] = 19; __Check((map.ContainsKey(18) && map[18] == 19).ToString(), "True");
