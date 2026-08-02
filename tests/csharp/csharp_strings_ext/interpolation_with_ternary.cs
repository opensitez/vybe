// vybe-test: csharp/csharp_strings_ext/interpolation_with_ternary
// origin: languages/csharp/tests/csharp/test_csharp_strings_ext.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
__Check(($"x is {(x > 3 ? "big" : "small")}").ToString(), "x is big");
