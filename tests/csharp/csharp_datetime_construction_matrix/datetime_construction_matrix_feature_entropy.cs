// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
string feature = "datetime_construction_matrix:94"; __Check((feature.Length >= 1).ToString(), "True");
