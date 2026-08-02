// vybe-test: csharp/csharp_string_methods/split_with_remove_empty_entries_drops_consecutive_delimiters
// origin: languages/csharp/tests/csharp/test_csharp_string_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var p = "a,,b".Split(new[]{','}, System.StringSplitOptions.RemoveEmptyEntries);
__Check((p.Length).ToString(), "2");
