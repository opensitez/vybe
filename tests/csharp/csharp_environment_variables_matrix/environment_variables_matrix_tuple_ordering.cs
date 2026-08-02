// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
var tuple = (left: 100, right: 101); __Check((tuple.left < tuple.right).ToString(), "True");
