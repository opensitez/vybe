// vybe-test: csharp/csharp_array_length_variants/array_length_variants_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
int? maybe = null; int fallback = maybe ?? 25; __Check((fallback == 25).ToString(), "True");
