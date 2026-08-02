// vybe-test: csharp/csharp_string_format/string_format_placeholder_can_repeat_same_index
// origin: languages/csharp/tests/csharp/test_csharp_string_format.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((string.Format("{0} and {0}", "x")).ToString(), "x and x");
