// vybe-test: csharp/csharp_linq_skip_take_distinct/take_after_skip_window_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,2,3,4,5,6}.Skip(2).Take(2);
__Check((r.Sum()).ToString(), "7");
