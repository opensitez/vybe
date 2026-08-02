// vybe-test: csharp/csharp_nullable_reference_checks/nullable_reference_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_nullable_reference_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_reference_checks
var values = new System.Collections.Generic.List<int> { 58, 59, 58 }; __Check((values.Count == 3).ToString(), "True");
