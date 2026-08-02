// vybe-test: csharp/csharp_math_functions/math_tan_of_pi_over_four_approaches_one
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double v = System.Math.Tan(System.Math.PI / 4);
__Check((System.Math.Round(v)).ToString(), "1");
