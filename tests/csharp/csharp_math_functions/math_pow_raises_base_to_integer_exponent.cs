// vybe-test: csharp/csharp_math_functions/math_pow_raises_base_to_integer_exponent
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Pow(2, 10)).ToString(), "1024");
