// vybe-test: csharp/csharp_struct_layout_matrix/struct_layout_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_struct_layout_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// struct_layout_matrix
string feature = "struct_layout_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
