// vybe-test: csharp/csharp_do_while_matrix/do_while_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_do_while_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// do_while_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[48] = 49; __Check((map.ContainsKey(48) && map[48] == 49).ToString(), "True");
