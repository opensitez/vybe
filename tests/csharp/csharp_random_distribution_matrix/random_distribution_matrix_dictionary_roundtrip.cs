// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[98] = 99; __Check((map.ContainsKey(98) && map[98] == 99).ToString(), "True");
