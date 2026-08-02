// vybe-test: csharp/csharp_array_length_variants/array_length_variants_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
int seed = 25; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
