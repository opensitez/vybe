// vybe-test: csharp/csharp_primary_constructors/primary_constructor_three_params_chained_sum
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Triple(int a, int b, int c) { public int Total => a + b + c; }
__Check((new Triple(1, 2, 3).Total).ToString(), "6");
