// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
string feature = "guid_parse_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
