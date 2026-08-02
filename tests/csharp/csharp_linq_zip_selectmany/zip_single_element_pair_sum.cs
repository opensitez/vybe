// vybe-test: csharp/csharp_linq_zip_selectmany/zip_single_element_pair_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{7}.Zip(new[]{5},(a,b)=>a+b);
__Check((z.Single()).ToString(), "12");
