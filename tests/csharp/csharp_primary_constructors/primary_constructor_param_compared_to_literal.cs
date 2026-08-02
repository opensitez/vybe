// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_compared_to_literal
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Check(int n) { public bool IsTen() => n == 10; }
__Check((new Check(10).IsTen()).ToString(), "True");
