// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
var values = new System.Collections.Generic.List<int> { 95, 96, 95 }; __Check((values.Count == 3).ToString(), "True");
