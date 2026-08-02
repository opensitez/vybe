// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
int seed = 57; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
