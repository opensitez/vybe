// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_while_none_skipped_returns_full_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{5,6,7}.SkipWhile(x=>x<3);
__Check((r.Count()).ToString(), "3");
