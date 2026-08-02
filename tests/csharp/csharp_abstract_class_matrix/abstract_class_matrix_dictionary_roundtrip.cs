// vybe-test: csharp/csharp_abstract_class_matrix/abstract_class_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_abstract_class_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// abstract_class_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[72] = 73; __Check((map.ContainsKey(72) && map[72] == 73).ToString(), "True");
