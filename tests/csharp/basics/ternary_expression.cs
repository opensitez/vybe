// vybe-test: csharp/basics/ternary_expression
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((5 > 3 ? "yes" : "no").ToString(), "yes");
