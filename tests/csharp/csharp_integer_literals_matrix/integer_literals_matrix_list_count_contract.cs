// vybe-test: csharp/csharp_integer_literals_matrix/integer_literals_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_integer_literals_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// integer_literals_matrix
var values = new System.Collections.Generic.List<int> { 15, 16, 15 }; __Check((values.Count == 3).ToString(), "True");
