// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[103] = 104; __Check((map.ContainsKey(103) && map[103] == 104).ToString(), "True");
