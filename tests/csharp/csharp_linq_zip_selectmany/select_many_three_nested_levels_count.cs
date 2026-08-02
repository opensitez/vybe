// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_three_nested_levels_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var data=new[]{new[]{new[]{1,2}},new[]{new[]{3}}};
var flat=data.SelectMany(a=>a).SelectMany(b=>b);
__Check((flat.Count()).ToString(), "3");
