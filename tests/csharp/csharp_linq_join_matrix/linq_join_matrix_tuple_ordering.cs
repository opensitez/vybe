// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
var tuple = (left: 119, right: 120); __Check((tuple.left < tuple.right).ToString(), "True");
