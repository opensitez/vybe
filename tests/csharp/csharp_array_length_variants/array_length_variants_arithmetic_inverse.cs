// vybe-test: csharp/csharp_array_length_variants/array_length_variants_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_length_variants
int seed = 25; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
