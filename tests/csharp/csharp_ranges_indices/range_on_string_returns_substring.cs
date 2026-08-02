// vybe-test: csharp/csharp_ranges_indices/range_on_string_returns_substring
// origin: languages/csharp/tests/csharp/test_csharp_ranges_indices.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string s="hello world"; __Check((s[6..]).ToString(), "world");
