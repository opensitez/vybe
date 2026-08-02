// vybe-test: csharp/csharp_linq_zip_selectmany/zip_char_sequences_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{'a','b'}.Zip(new[]{'x','y'},(a,b)=>(char)(a+b-'a'+'x'));
__Check((z.Count()).ToString(), "2");
