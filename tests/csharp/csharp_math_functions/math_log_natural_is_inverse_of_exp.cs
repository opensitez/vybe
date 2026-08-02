// vybe-test: csharp/csharp_math_functions/math_log_natural_is_inverse_of_exp
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double v = System.Math.Log(System.Math.E);
__Check((System.Math.Round(v, 5)).ToString(), "1");
