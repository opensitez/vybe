// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_all_same_key_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{10,20,30}.DistinctBy(n=>0);
__Check((r.Count()).ToString(), "1");
