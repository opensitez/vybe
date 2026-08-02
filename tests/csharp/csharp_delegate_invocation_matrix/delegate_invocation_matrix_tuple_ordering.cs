// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
var tuple = (left: 74, right: 75); __Check((tuple.left < tuple.right).ToString(), "True");
