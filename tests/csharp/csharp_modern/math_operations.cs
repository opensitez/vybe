// vybe-test: csharp/csharp_modern/math_operations
// origin: languages/csharp/tests/csharp/test_csharp_modern.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Math.Abs(-42)).ToString(), "42");
__Check((Math.Max(10, 20)).ToString(), "20");
__Check((Math.Min(10, 20)).ToString(), "10");
__Check((Math.Sqrt(25)).ToString(), "5");
__Check((Math.Floor(3.7)).ToString(), "3");
__Check((Math.Ceiling(3.2)).ToString(), "4");
