// vybe-test: csharp/math/math_max_selects_larger_of_two_doubles
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.Math.Max(1.5, 2.5)).ToString(), "2.5");
