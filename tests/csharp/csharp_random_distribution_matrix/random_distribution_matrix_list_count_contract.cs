// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// random_distribution_matrix
var values = new System.Collections.Generic.List<int> { 98, 99, 98 }; __Check((values.Count == 3).ToString(), "True");
