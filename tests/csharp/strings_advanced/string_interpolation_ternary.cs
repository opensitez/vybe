// vybe-test: csharp/strings_advanced/string_interpolation_ternary
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 10;
__Check(($"x is {(x > 5 ? "big" : "small")}").ToString(), "x is big");
