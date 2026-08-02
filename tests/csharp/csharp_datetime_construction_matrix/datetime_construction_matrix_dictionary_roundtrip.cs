// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[94] = 95; __Check((map.ContainsKey(94) && map[94] == 95).ToString(), "True");
