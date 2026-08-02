// vybe-test: csharp/csharp_array_length_variants/array_length_variants_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
string feature = "array_length_variants"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
