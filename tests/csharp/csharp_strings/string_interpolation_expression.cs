// vybe-test: csharp/csharp_strings/string_interpolation_expression
// origin: languages/csharp/tests/csharp/test_csharp_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a = 3, b = 4;
__Check(($"sum = {a + b}").ToString(), "sum = 7");
