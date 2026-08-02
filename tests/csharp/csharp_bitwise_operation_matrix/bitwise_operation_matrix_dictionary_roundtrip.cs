// vybe-test: csharp/csharp_bitwise_operation_matrix/bitwise_operation_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_bitwise_operation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// bitwise_operation_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[104] = 105; __Check((map.ContainsKey(104) && map[104] == 105).ToString(), "True");
