// vybe-test: csharp/csharp_io_stream_matrix/io_stream_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_io_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_stream_matrix
var values = new System.Collections.Generic.List<int> { 90, 91, 90 }; __Check((values.Count == 3).ToString(), "True");
