// vybe-test: csharp/csharp_io_path_matrix/io_path_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_io_path_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_path_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[122] = 123; __Check((map.ContainsKey(122) && map[122] == 123).ToString(), "True");
