// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
int? maybe = 81; __Check((maybe.HasValue && maybe.Value == 81).ToString(), "True");
