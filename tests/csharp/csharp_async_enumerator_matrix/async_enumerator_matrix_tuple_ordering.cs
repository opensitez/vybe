// vybe-test: csharp/csharp_async_enumerator_matrix/async_enumerator_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_async_enumerator_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_enumerator_matrix
var tuple = (left: 116, right: 117); __Check((tuple.left < tuple.right).ToString(), "True");
