// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
var values = new System.Collections.Generic.List<int> { 112, 113, 112 }; __Check((values.Count == 3).ToString(), "True");
