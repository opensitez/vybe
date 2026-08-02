// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
var tuple = (left: 111, right: 112); __Check((tuple.left < tuple.right).ToString(), "True");
