// vybe-test: csharp/csharp_linq_zip_selectmany/zip_after_skip_on_first_side_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{1,2,3,4}.Skip(1).Zip(new[]{10,20,30},(a,b)=>a+b);
__Check((z.Count()).ToString(), "3");
