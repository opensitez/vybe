// vybe-test: csharp/csharp_primary_constructors/primary_constructor_long_param_used
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Big(long n) { public long Value => n; }
__Check((new Big(9000000000L).Value).ToString(), "9000000000");
