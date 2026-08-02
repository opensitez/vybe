// vybe-test: csharp/csharp_decimal_math_matrix/decimal_math_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_decimal_math_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// decimal_math_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[17] = 18; __Check((map.ContainsKey(17) && map[17] == 18).ToString(), "True");
