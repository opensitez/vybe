// vybe-test: csharp/csharp_datetime_format_matrix/datetime_format_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_format_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_format_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[96] = 97; __Check((map.ContainsKey(96) && map[96] == 97).ToString(), "True");
