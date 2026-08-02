// vybe-test: csharp/csharp_environment_variables_matrix/environment_variables_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_environment_variables_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// environment_variables_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[100] = 101; __Check((map.ContainsKey(100) && map[100] == 101).ToString(), "True");
