// vybe-test: csharp/strings_advanced/fully_qualified_system_string_format
// origin: languages/csharp/tests/csharp/test_strings_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((System.String.Format("{0}-{1}", "A", "B")).ToString(), "A-B");
