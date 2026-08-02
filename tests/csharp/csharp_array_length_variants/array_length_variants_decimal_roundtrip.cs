// vybe-test: csharp/csharp_array_length_variants/array_length_variants_decimal_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
decimal amount = 10m; __Check(((amount / 2m) * 2m == 10m).ToString(), "True");
