// vybe-test: csharp/csharp_primary_constructors/primary_constructor_short_param_doubled
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class ShortScale(short n) { public int Twice => n * 2; }
__Check((new ShortScale(50).Twice).ToString(), "100");
