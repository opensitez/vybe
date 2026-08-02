// vybe-test: csharp/csharp_math_trigonometry_matrix/math_trigonometry_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_math_trigonometry_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_trigonometry_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[102] = 103; __Check((map.ContainsKey(102) && map[102] == 103).ToString(), "True");
