// vybe-test: csharp/csharp_math_advanced/math_floor_rounds_toward_negative_infinity
// origin: languages/csharp/tests/csharp/test_csharp_math_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Floor(2.9)).ToString(), "2");
