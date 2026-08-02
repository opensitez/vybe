// vybe-test: csharp/csharp_datetime_construction_matrix/datetime_construction_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_datetime_construction_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// datetime_construction_matrix
string feature = "datetime_construction_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
