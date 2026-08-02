// vybe-test: csharp/common_patterns/math_portable_stdlib_functions
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Sin(0)).ToString(), "0");
__Check((Math.Cos(0)).ToString(), "1");
__Check((Math.Tan(0)).ToString(), "0");
__Check((Math.Asin(0)).ToString(), "0");
__Check((Math.Acos(1)).ToString(), "0");
__Check((Math.Atan(0)).ToString(), "0");
__Check((Math.Atan2(0, 1)).ToString(), "0");
__Check((Math.Log(1)).ToString(), "0");
__Check((Math.Log10(100)).ToString(), "2");
__Check((Math.Exp(0)).ToString(), "1");
__Check((Math.Sign(-5)).ToString(), "-1");
__Check((Math.Clamp(15, 0, 10)).ToString(), "10");
