// vybe-test: csharp/csharp_math_advanced/math_ceiling_rounds_toward_positive_infinity
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Ceiling(2.1)).ToString(), "3");
