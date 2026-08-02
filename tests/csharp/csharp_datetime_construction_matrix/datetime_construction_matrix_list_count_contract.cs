// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
var values = new System.Collections.Generic.List<int> { 94, 95, 94 }; __Check((values.Count == 3).ToString(), "True");
