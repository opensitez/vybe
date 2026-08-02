// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_strings_skip_prefix_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{"a","b","c","d"}.Skip(2);
__Check((r.Count()).ToString(), "2");
