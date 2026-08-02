// vybe-test: csharp/strings/string_split
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts = "a,b,c".Split(",");
        __Check((parts[0]).ToString(), "a");
        __Check((parts[1]).ToString(), "b");
        __Check((parts[2]).ToString(), "c");
