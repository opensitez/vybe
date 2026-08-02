// vybe-test: csharp/csharp_cast_runtime_matrix/cast_runtime_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_cast_runtime_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// cast_runtime_matrix
var tuple = (left: 61, right: 62); __Check((tuple.left < tuple.right).ToString(), "True");
