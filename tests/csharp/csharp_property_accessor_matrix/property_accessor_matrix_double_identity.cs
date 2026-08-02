// vybe-test: csharp/csharp_property_accessor_matrix/property_accessor_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_property_accessor_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// property_accessor_matrix
double seed = 64; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
