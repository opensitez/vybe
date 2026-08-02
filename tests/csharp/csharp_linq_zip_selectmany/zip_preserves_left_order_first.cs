// vybe-test: csharp/csharp_linq_zip_selectmany/zip_preserves_left_order_first
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{3,1,2}.Zip(new[]{1,1,1},(a,b)=>a);
__Check((z.First()).ToString(), "3");
