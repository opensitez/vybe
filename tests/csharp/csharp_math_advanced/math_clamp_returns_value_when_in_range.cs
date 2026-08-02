// vybe-test: csharp/csharp_math_advanced/math_clamp_returns_value_when_in_range
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Clamp(5,0,10)).ToString(), "5");
