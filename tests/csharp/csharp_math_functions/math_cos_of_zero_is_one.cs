// vybe-test: csharp/csharp_math_functions/math_cos_of_zero_is_one
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double v = System.Math.Cos(0);
__Check((v).ToString(), "1");
