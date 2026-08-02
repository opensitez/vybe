// vybe-test: csharp/csharp_method_group_matrix/method_group_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_method_group_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// method_group_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[79] = 80; __Check((map.ContainsKey(79) && map[79] == 80).ToString(), "True");
