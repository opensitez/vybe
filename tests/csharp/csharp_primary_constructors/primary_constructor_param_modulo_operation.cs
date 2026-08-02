// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_modulo_operation
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Mod(int n) { public int Rem(int d) => n % d; }
__Check((new Mod(10).Rem(3)).ToString(), "1");
