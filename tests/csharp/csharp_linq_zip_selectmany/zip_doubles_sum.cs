// vybe-test: csharp/csharp_linq_zip_selectmany/zip_doubles_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{1.5,2.5}.Zip(new[]{2.0,2.0},(a,b)=>a+b);
__Check((z.Sum()).ToString(), "8");
