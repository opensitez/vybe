// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[53] = 54; __Check((map.ContainsKey(53) && map[53] == 54).ToString(), "True");
