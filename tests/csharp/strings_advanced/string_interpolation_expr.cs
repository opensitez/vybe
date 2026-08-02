// vybe-test: csharp/strings_advanced/string_interpolation_expr
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
__Check(($"{x} squared is {x * x}").ToString(), "5 squared is 25");
