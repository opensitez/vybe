// vybe-test: csharp/common_patterns/math_min_max_round
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Min(3, 7)).ToString(), "3");
__Check((Math.Max(3, 7)).ToString(), "7");
__Check((Math.Round(3.7)).ToString(), "4");
__Check((Math.Floor(3.7)).ToString(), "3");
__Check((Math.Ceiling(3.2)).ToString(), "4");
