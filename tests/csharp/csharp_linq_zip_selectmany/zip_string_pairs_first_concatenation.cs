// vybe-test: csharp/csharp_linq_zip_selectmany/zip_string_pairs_first_concatenation
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{"a","b"}.Zip(new[]{"1","2"},(x,y)=>x+y);
__Check((z.First()).ToString(), "a1");
