// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[121] = 122; __Check((map.ContainsKey(121) && map[121] == 122).ToString(), "True");
