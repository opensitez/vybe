// vybe-test: csharp/csharp_linq_skip_take_distinct/take_one_single_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_skip_take_distinct.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var r=new[]{99,100,101}.Take(1);
__Check((r.Single()).ToString(), "99");
