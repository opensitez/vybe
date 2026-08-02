// vybe-test: csharp/csharp_attribute_visibility_matrix/attribute_visibility_matrix_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_attribute_visibility_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// attribute_visibility_matrix
int seed = 93; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
