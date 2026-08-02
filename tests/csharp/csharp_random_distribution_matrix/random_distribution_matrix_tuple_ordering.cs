// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
var tuple = (left: 98, right: 99); __Check((tuple.left < tuple.right).ToString(), "True");
