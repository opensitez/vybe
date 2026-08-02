// vybe-test: csharp/csharp_generic_constraints_matrix/generic_constraints_matrix_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_constraints_matrix
string feature = "generic_constraints_matrix"; __Check((feature.Contains("a") || !feature.Contains("a")).ToString(), "True");
