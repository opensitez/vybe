// vybe-test: csharp/csharp_io_stream_matrix/io_stream_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_io_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_stream_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(90); set.Add(90); __Check((set.Count == 1).ToString(), "True");
