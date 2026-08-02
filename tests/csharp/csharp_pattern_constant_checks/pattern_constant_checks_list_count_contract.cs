// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pattern_constant_checks
var values = new System.Collections.Generic.List<int> { 40, 41, 40 }; __Check((values.Count == 3).ToString(), "True");
