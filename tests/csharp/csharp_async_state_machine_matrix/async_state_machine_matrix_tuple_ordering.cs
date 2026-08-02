// vybe-test: csharp/csharp_async_state_machine_matrix/async_state_machine_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_async_state_machine_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_state_machine_matrix
var tuple = (left: 88, right: 89); __Check((tuple.left < tuple.right).ToString(), "True");
