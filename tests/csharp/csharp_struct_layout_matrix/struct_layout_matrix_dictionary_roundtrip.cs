// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[113] = 114; __Check((map.ContainsKey(113) && map[113] == 114).ToString(), "True");
