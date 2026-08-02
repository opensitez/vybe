// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_format_matrix
var values = new System.Collections.Generic.List<int> { 96, 97, 96 }; __Check((values.Count == 3).ToString(), "True");
