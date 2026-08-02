// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_expand_each_to_range_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{1,2,3}.SelectMany(n=>Enumerable.Range(1,n));
__Check((flat.Sum()).ToString(), "10");
