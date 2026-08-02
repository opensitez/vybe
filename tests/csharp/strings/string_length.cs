// vybe-test: csharp/strings/string_length
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s = "hello";
        __Check((s.Substring(0, 3)).ToString(), "hel");
