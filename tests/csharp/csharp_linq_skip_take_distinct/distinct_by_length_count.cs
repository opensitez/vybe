// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_length_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{"a","bb","c","dd","eee"}.DistinctBy(s=>s.Length);
__Check((r.Count()).ToString(), "3");
