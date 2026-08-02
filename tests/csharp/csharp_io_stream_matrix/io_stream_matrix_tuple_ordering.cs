// vybe-test: csharp/csharp_io_stream_matrix/io_stream_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_io_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_stream_matrix
var tuple = (left: 90, right: 91); __Check((tuple.left < tuple.right).ToString(), "True");
