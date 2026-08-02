// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
var tuple = (left: 80, right: 81); __Check((tuple.left < tuple.right).ToString(), "True");
