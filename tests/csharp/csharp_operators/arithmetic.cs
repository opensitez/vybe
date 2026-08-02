// vybe-test: csharp/csharp_operators/arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((10 + 5).ToString(), "15");
__Check((10 - 5).ToString(), "5");
__Check((10 * 5).ToString(), "50");
__Check((10 % 3).ToString(), "1");
