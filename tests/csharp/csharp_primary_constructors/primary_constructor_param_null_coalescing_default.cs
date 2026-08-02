// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_null_coalescing_default
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Maybe(string? text) { public string Safe() => text ?? "none"; }
__Check((new Maybe(null).Safe()).ToString(), "none");
