// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_arithmetic_inverse
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
int seed = 29; __Check(((seed * 2) / 2 == seed || seed == 0).ToString(), "True");
