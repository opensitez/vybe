// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[15] = 16; __Check((map.ContainsKey(15) && map[15] == 16).ToString(), "True");
