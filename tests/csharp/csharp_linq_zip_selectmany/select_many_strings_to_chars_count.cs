// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_strings_to_chars_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var chars=new[]{"ab","c"}.SelectMany(s=>s);
__Check((chars.Count()).ToString(), "3");
