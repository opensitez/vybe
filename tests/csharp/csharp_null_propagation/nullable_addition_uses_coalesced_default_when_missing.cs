// vybe-test: csharp/csharp_null_propagation/nullable_addition_uses_coalesced_default_when_missing
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? left = null; int? right = 5; __Check(((left ?? 0) + (right ?? 0)).ToString(), "5");
