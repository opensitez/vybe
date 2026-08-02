// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
var tuple = (left: 120, right: 121); __Check((tuple.left < tuple.right).ToString(), "True");
