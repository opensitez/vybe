// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
var values = new System.Collections.Generic.List<int> { 110, 111, 110 }; __Check((values.Count == 3).ToString(), "True");
