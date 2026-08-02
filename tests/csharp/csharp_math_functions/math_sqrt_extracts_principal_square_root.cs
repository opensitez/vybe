// vybe-test: csharp/csharp_math_functions/math_sqrt_extracts_principal_square_root
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Sqrt(144)).ToString(), "12");
