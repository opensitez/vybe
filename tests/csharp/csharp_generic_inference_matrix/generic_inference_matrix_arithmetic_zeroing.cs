// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
int seed = 81; __Check((seed - seed == 0).ToString(), "True");
