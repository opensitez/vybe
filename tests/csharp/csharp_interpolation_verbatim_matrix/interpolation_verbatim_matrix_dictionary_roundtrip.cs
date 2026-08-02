// vybe-test: csharp/csharp_interpolation_verbatim_matrix/interpolation_verbatim_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_verbatim_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_verbatim_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[110] = 111; __Check((map.ContainsKey(110) && map[110] == 111).ToString(), "True");
