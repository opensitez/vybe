// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
var values = new System.Collections.Generic.List<int> { 36, 37, 36 }; __Check((values.Count == 3).ToString(), "True");
