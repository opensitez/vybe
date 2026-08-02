// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[58] = 59; __Check((map.ContainsKey(58) && map[58] == 59).ToString(), "True");
