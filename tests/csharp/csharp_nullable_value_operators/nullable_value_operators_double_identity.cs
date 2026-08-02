// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
double seed = 57; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
