// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_flatten_jagged_arrays_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{new[]{1,2},new[]{3,4,5}}.SelectMany(x=>x);
__Check((flat.Count()).ToString(), "5");
