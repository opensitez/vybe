// vybe-test: csharp/csharp_operators/parenthesized_precedence
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(((2 + 3) * 4).ToString(), "20");
__Check((2 + 3 * 4).ToString(), "14");
