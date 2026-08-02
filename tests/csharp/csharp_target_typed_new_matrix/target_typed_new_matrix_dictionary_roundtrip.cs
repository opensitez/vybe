// vybe-test: csharp/csharp_target_typed_new_matrix/target_typed_new_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// target_typed_new_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[107] = 108; __Check((map.ContainsKey(107) && map[107] == 108).ToString(), "True");
