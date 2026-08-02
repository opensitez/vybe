// vybe-test: csharp/math/math_round_midpoint_to_even_for_half_values
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Round(2.5)).ToString(), "2");
