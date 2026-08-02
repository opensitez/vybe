// vybe-test: csharp/csharp_linq_advanced/skip_last_omits_trailing_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var result=new[]{1,2,3,4,5}.SkipLast(2);
__Check((result.Count()).ToString(), "3");
