// vybe-test: csharp/csharp_regex_capture_groups/regex_split_returns_segments_around_delimiter_pattern
// origin: languages/csharp/tests/csharp/test_csharp_regex_capture_groups.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var parts = System.Text.RegularExpressions.Regex.Split("one,two,three", ",");
__Check((parts[1]).ToString(), "two");
