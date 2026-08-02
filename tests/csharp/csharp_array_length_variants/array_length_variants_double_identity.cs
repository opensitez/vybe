// vybe-test: csharp/csharp_array_length_variants/array_length_variants_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
double seed = 25; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
