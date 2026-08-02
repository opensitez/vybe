// vybe-test: csharp/strings_advanced/string_format
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0} + {1} = {2}", 1, 2, 3)).ToString(), "1 + 2 = 3");
__Check((string.Format("Name: {0}, Age: {1}", "Bob", 25)).ToString(), "Name: Bob, Age: 25");
