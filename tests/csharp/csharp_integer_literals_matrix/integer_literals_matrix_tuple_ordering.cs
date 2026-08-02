// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
var tuple = (left: 15, right: 16); __Check((tuple.left < tuple.right).ToString(), "True");
