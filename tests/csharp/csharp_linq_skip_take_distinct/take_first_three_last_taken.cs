// vybe-test: csharp/csharp_linq_skip_take_distinct/take_first_three_last_taken
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{10,20,30,40,50}.Take(3);
__Check((r.Last()).ToString(), "30");
