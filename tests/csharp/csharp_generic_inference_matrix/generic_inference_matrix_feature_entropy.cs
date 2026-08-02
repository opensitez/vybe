// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
string feature = "generic_inference_matrix:81"; __Check((feature.Length >= 1).ToString(), "True");
