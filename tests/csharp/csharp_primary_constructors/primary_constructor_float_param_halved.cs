// vybe-test: csharp/csharp_primary_constructors/primary_constructor_float_param_halved
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Half(float n) { public float Value => n / 2f; }
__Check((new Half(10f).Value).ToString(), "5");
