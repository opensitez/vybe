// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
string feature = "multidim_array_traversal"; __Check((feature.Length > 0).ToString(), "True");
