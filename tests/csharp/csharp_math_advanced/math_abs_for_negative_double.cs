// vybe-test: csharp/csharp_math_advanced/math_abs_for_negative_double
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Abs(-3.7)).ToString(), "3.7");
