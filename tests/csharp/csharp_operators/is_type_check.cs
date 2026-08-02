// vybe-test: csharp/csharp_operators/is_type_check
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object x = "hello";
__Check((x is string).ToString(), "True");
