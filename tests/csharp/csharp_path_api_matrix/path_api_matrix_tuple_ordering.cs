// vybe-test: csharp/csharp_path_api_matrix/path_api_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_path_api_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// path_api_matrix
var tuple = (left: 123, right: 124); __Check((tuple.left < tuple.right).ToString(), "True");
