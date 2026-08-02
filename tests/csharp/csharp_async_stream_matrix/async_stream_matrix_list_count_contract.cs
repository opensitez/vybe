// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
var values = new System.Collections.Generic.List<int> { 111, 112, 111 }; __Check((values.Count == 3).ToString(), "True");
