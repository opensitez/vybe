// vybe-test: csharp/csharp_string_immutability_checks/string_immutability_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_string_immutability_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_immutability_checks
var values = new System.Collections.Generic.List<int> { 18, 19, 18 }; __Check((values.Count == 3).ToString(), "True");
