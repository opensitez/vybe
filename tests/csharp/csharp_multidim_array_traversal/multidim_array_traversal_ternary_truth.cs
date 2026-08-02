// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
int seed = 29; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
