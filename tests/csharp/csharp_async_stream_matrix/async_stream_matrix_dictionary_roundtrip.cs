// vybe-test: csharp/csharp_async_stream_matrix/async_stream_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_async_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// async_stream_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[111] = 112; __Check((map.ContainsKey(111) && map[111] == 112).ToString(), "True");
