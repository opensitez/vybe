// vybe-test: csharp/math/math_sqrt
// origin: languages/csharp/tests/csharp/test_math.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Sqrt(16)).ToString(), "4");
