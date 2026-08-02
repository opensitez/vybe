// vybe-test: csharp/csharp_interlocked_atomic_matrix/interlocked_atomic_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_interlocked_atomic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interlocked_atomic_matrix
var tuple = (left: 83, right: 84); __Check((tuple.left < tuple.right).ToString(), "True");
