// vybe-test: csharp/csharp_operator_overloading_matrix/operator_overloading_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// operator_overloading_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[105] = 106; __Check((map.ContainsKey(105) && map[105] == 106).ToString(), "True");
