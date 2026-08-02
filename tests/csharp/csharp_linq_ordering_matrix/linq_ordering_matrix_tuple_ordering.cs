// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
var tuple = (left: 121, right: 122); __Check((tuple.left < tuple.right).ToString(), "True");
