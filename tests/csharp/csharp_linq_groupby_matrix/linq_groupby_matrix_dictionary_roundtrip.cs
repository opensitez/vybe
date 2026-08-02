// vybe-test: csharp/csharp_linq_groupby_matrix/linq_groupby_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_groupby_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_groupby_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[120] = 121; __Check((map.ContainsKey(120) && map[120] == 121).ToString(), "True");
