// vybe-test: csharp/csharp_linq_skip_take_distinct/distinct_by_on_empty_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=System.Array.Empty<int>().DistinctBy(n=>n);
__Check((r.Count()).ToString(), "0");
