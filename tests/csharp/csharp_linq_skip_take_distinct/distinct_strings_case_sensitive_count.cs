// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_strings_case_sensitive_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{"A","a","A","b"}.Distinct();
__Check((r.Count()).ToString(), "3");
