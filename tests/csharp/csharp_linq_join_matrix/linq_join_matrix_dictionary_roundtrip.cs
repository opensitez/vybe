// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[119] = 120; __Check((map.ContainsKey(119) && map[119] == 120).ToString(), "True");
