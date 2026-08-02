// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
string feature = "multidim_array_traversal:29"; __Check((feature.Length >= 1).ToString(), "True");
