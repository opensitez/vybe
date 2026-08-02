// vybe-test: csharp/csharp_threading_pool_matrix/threading_pool_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_threading_pool_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// threading_pool_matrix
var tuple = (left: 87, right: 88); __Check((tuple.left < tuple.right).ToString(), "True");
