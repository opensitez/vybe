// vybe-test: csharp/csharp_linq_skip_take_distinct/take_while_all_match_returns_full_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{2,4,6}.TakeWhile(x=>x%2==0);
__Check((r.Count()).ToString(), "3");
