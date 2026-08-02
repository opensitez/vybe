// vybe-test: csharp/csharp_linq_zip_selectmany/zip_result_selector_returns_tuple_like_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{1,2}.Zip(new[]{3,4},(a,b)=>a*10+b);
__Check((z.Sum()).ToString(), "47");
