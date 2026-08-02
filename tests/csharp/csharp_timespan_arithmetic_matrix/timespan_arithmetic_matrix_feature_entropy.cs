// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
string feature = "timespan_arithmetic_matrix:95"; __Check((feature.Length >= 1).ToString(), "True");
