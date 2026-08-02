// vybe-test: csharp/csharp_array_length_variants/array_length_variants_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
string feature = "array_length_variants:25"; __Check((feature.Length >= 1).ToString(), "True");
