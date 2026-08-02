// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
string feature = "generic_inference_matrix"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
