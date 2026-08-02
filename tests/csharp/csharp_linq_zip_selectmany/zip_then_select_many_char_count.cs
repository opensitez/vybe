// vybe-test: csharp/csharp_linq_zip_selectmany/zip_then_select_many_char_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var words=new[]{"hi","go"};
var letters=words.Zip(new[]{1,2},(w,n)=>w).SelectMany(w=>w);
__Check((letters.Count()).ToString(), "4");
