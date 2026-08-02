// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
var values = new System.Collections.Generic.List<int> { 99, 100, 99 }; __Check((values.Count == 3).ToString(), "True");
