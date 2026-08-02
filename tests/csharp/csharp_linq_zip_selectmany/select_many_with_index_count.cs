// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_with_index_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{new[]{10},new[]{20,30}}.SelectMany((x,i)=>x.Select(v=>v+i));
__Check((flat.Count()).ToString(), "3");
