// vybe-test: csharp/csharp_goto_label_matrix/goto_label_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_goto_label_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// goto_label_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[50] = 51; __Check((map.ContainsKey(50) && map[50] == 51).ToString(), "True");
