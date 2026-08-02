// vybe-test: csharp/csharp_linq_skip_take_distinct/take_more_than_length_returns_all_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,2}.Take(10);
__Check((r.Count()).ToString(), "2");
