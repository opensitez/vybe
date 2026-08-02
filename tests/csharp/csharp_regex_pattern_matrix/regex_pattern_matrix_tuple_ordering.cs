// vybe-test: csharp/csharp_regex_pattern_matrix/regex_pattern_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_regex_pattern_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// regex_pattern_matrix
var tuple = (left: 99, right: 100); __Check((tuple.left < tuple.right).ToString(), "True");
