// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
double seed = 81; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
