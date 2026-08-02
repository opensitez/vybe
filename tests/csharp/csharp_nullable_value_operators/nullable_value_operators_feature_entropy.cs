// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
string feature = "nullable_value_operators:57"; __Check((feature.Length >= 1).ToString(), "True");
