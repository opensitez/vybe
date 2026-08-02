// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
int seed = 64; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
