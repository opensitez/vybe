// vybe-test: csharp/csharp_pointer_like_emulation_matrix/pointer_like_emulation_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_pointer_like_emulation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// pointer_like_emulation_matrix
string feature = "pointer_like_emulation_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
