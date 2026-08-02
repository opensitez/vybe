// vybe-test: csharp/csharp_operators/logical_operators
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((true && true).ToString(), "True");
__Check((true && false).ToString(), "False");
__Check((false || true).ToString(), "True");
__Check((false || false).ToString(), "False");
__Check((!true).ToString(), "False");
