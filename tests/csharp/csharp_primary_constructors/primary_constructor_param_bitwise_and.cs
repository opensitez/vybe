// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_bitwise_and
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Mask(int n) { public int And(int m) => n & m; }
__Check((new Mask(12).And(10)).ToString(), "8");
