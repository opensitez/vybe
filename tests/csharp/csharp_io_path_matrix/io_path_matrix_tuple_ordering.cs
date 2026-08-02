// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
var tuple = (left: 122, right: 123); __Check((tuple.left < tuple.right).ToString(), "True");
