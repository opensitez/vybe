// vybe-test: csharp/csharp_array_indexing_matrix/array_indexing_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_array_indexing_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_indexing_matrix
var tuple = (left: 24, right: 25); __Check((tuple.left < tuple.right).ToString(), "True");
