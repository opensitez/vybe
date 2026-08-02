// vybe-test: csharp/csharp_math_functions/math_sign_returns_minus_one_for_negative
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Sign(-42)).ToString(), "-1");
