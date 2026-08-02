// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
int seed = 97; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
