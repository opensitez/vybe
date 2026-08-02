// vybe-test: csharp/csharp_linq_skip_take_distinct/skip_last_element_via_skip_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{1,2,3,4}.Skip(3);
__Check((r.Count()).ToString(), "1");
