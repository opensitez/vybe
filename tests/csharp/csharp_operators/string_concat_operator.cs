// vybe-test: csharp/csharp_operators/string_concat_operator
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check(("Hello" + " " + "World").ToString(), "Hello World");
__Check(("num: " + 42).ToString(), "num: 42");
