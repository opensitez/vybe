// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_doubles_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{new[]{1.5,2.5},new[]{3.0}}.SelectMany(x=>x);
__Check((flat.Count()).ToString(), "3");
