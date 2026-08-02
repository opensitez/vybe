// vybe-test: csharp/math/math_clamp_restricts_value_to_inclusive_bounds
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Clamp(10, 0, 5)).ToString(), "5");
