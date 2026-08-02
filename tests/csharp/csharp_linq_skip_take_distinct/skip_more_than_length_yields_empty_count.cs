// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_more_than_length_yields_empty_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,2,3}.Skip(10);
__Check((r.Count()).ToString(), "0");
