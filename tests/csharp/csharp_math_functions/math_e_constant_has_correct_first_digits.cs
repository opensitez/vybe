// vybe-test: csharp/csharp_math_functions/math_e_constant_has_correct_first_digits
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.E > 2.71 && System.Math.E < 2.72).ToString(), "True");
