// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_while_strings_skip_short_prefix_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{"a","bb","ccc","d"}.SkipWhile(s=>s.Length<3);
__Check((r.Count()).ToString(), "2");
