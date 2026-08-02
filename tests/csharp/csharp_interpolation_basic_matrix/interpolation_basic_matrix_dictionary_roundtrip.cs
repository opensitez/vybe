// vybe-test: csharp/csharp_interpolation_basic_matrix/interpolation_basic_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_interpolation_basic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// interpolation_basic_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[112] = 113; __Check((map.ContainsKey(112) && map[112] == 113).ToString(), "True");
