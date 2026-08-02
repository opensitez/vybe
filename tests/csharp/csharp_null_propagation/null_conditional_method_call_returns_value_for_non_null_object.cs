// vybe-test: csharp/csharp_null_propagation/null_conditional_method_call_returns_value_for_non_null_object
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string text = "hello"; __Check((text?.ToUpper()).ToString(), "HELLO");
