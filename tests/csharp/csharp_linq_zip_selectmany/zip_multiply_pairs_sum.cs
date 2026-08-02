// vybe-test: csharp/csharp_linq_zip_selectmany/zip_multiply_pairs_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{1,2,3}.Zip(new[]{10,20,30},(a,b)=>a*b);
__Check((z.Sum()).ToString(), "140");
