// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_in_conditional_expression
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Sign(int n) { public string Kind() => n >= 0 ? "pos" : "neg"; }
__Check((new Sign(5).Kind()).ToString(), "pos");
