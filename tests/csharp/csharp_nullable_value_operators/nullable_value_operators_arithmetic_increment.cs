// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
int seed = 57; __Check((seed + 1 > seed).ToString(), "True");
