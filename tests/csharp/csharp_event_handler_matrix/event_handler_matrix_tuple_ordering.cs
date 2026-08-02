// vybe-test: csharp/csharp_event_handler_matrix/event_handler_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_event_handler_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// event_handler_matrix
var tuple = (left: 77, right: 78); __Check((tuple.left < tuple.right).ToString(), "True");
