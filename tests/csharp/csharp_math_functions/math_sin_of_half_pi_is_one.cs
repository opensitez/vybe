// vybe-test: csharp/csharp_math_functions/math_sin_of_half_pi_is_one
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double v = System.Math.Sin(System.Math.PI / 2);
__Check((System.Math.Round(v)).ToString(), "1");
