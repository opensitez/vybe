// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_repeat_each_element_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{1,2}.SelectMany(n=>new[]{n,n,n});
__Check((flat.Sum()).ToString(), "9");
