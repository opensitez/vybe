// vybe-test: csharp/csharp_linq_zip_selectmany/zip_negative_numbers_product_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{-1,2}.Zip(new[]{3,-4},(a,b)=>a*b);
__Check((z.Count()).ToString(), "2");
