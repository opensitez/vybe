// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_shift_left
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Shift(int n) { public int Left(int bits) => n << bits; }
__Check((new Shift(3).Left(2)).ToString(), "12");
