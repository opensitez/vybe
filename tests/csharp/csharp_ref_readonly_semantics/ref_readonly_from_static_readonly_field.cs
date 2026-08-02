// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_from_static_readonly_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Defaults{public static readonly int Seed=15;} ref readonly int SeedRef()=>ref Defaults.Seed; __Check((SeedRef()).ToString(), "15");
