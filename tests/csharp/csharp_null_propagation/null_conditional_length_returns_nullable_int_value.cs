// vybe-test: csharp/csharp_null_propagation/null_conditional_length_returns_nullable_int_value
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text = "four"; __Check((text?.Length ?? 0).ToString(), "4");
