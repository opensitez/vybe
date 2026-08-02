// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_then_order_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{"zzz","a","bb","c","dd"}.DistinctBy(s=>s.Length).OrderBy(s=>s);
__Check((r.Count()).ToString(), "3");
