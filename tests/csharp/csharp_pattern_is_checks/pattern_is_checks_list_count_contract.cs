// vybe-test: csharp/csharp_pattern_is_checks/pattern_is_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_pattern_is_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_is_checks
var values = new System.Collections.Generic.List<int> { 41, 42, 41 }; __Check((values.Count == 3).ToString(), "True");
