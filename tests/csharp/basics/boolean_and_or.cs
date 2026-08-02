// vybe-test: csharp/basics/boolean_and_or
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((true && false).ToString(), "False");
        __Check((true || false).ToString(), "True");
        __Check((!true).ToString(), "False");
