// vybe-test: csharp/common_patterns/enum_tostring_parse
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Day { Mon, Tue, Wed, Thu, Fri }
Day d = Day.Wed;
__Check((d.ToString()).ToString(), "Wed");
