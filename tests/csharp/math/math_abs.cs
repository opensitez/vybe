// vybe-test: csharp/math/math_abs
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Abs(-5)).ToString(), "5");
