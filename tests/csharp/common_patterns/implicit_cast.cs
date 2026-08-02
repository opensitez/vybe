// vybe-test: csharp/common_patterns/implicit_cast
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int i = 42;
double d = i;
__Check((d).ToString(), "42");
