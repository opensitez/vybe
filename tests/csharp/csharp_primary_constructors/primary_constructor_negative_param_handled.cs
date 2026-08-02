// vybe-test: csharp/csharp_primary_constructors/primary_constructor_negative_param_handled
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Sign(int n) { public int Abs() => n < 0 ? -n : n; }
__Check((new Sign(-8).Abs()).ToString(), "8");
