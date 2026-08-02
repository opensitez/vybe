// vybe-test: csharp/common_patterns/math_abs_pow_sqrt
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Abs(-42)).ToString(), "42");
__Check((Math.Pow(2, 10)).ToString(), "1024");
__Check((Math.Sqrt(144)).ToString(), "12");
