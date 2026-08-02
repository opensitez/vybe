// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_zip_then_select_many_sum
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pairs=new[]{1,2}.Zip(new[]{3,4},(a,b)=>new[]{a,b});
__Check((pairs.SelectMany(x=>x).Sum()).ToString(), "10");
