// vybe-test: csharp/csharp_pattern_positional_checks/pattern_positional_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_pattern_positional_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_positional_checks
var values = new System.Collections.Generic.List<int> { 115, 116, 115 }; __Check((values.Count == 3).ToString(), "True");
