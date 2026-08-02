// vybe-test: csharp/csharp_string_join_split/split_removes_empty_entries_option
// origin: languages/csharp/tests/csharp/test_csharp_string_join_split.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts="a,,b,,c".Split(',',System.StringSplitOptions.RemoveEmptyEntries);
__Check((parts.Length).ToString(), "3");
