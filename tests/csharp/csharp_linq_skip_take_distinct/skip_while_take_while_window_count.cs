// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_while_take_while_window_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,2,3,4,5,6,7}.SkipWhile(x=>x<3).TakeWhile(x=>x<6);
__Check((r.Count()).ToString(), "3");
