// vybe-test: csharp/csharp_math_minmax_matrix/math_minmax_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_math_minmax_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_minmax_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[101] = 102; __Check((map.ContainsKey(101) && map[101] == 102).ToString(), "True");
