// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_then_distinct_by_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,1,2,2,3,3}.Skip(2).DistinctBy(n=>n);
__Check((r.Count()).ToString(), "2");
