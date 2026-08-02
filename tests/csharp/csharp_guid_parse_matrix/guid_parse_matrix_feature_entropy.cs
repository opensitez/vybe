// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
string feature = "guid_parse_matrix:97"; __Check((feature.Length >= 1).ToString(), "True");
