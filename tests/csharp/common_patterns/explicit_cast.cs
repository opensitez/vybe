// vybe-test: csharp/common_patterns/explicit_cast
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double d = 3.99;
int i = (int)d;
__Check((i).ToString(), "3");
