// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
string feature = "nullable_value_operators"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
