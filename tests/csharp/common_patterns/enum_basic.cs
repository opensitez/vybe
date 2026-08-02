// vybe-test: csharp/common_patterns/enum_basic
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green, Blue }
Color c = Color.Green;
__Check((c).ToString(), "Green");
__Check(((int)c).ToString(), "1");
