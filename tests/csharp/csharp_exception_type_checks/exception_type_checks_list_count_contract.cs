// vybe-test: csharp/csharp_exception_type_checks/exception_type_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_exception_type_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// exception_type_checks
var values = new System.Collections.Generic.List<int> { 53, 54, 53 }; __Check((values.Count == 3).ToString(), "True");
