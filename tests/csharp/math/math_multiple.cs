// vybe-test: csharp/math/math_multiple
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Floor(9.7)).ToString(), "9");
        __Check((Math.Abs(-42)).ToString(), "42");
        __Check((Math.Sqrt(144)).ToString(), "12");
