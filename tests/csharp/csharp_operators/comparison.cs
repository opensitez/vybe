// vybe-test: csharp/csharp_operators/comparison
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((1 < 2).ToString(), "True");
__Check((2 > 1).ToString(), "True");
__Check((1 <= 1).ToString(), "True");
__Check((1 >= 1).ToString(), "True");
__Check((1 == 1).ToString(), "True");
__Check((1 != 2).ToString(), "True");
