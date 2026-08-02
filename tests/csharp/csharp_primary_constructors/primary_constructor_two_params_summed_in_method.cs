// vybe-test: csharp/csharp_primary_constructors/primary_constructor_two_params_summed_in_method
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Adder(int a, int b) { public int Sum() => a + b; }
__Check((new Adder(3, 4).Sum()).ToString(), "7");
