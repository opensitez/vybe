// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_empty_inner_sequences_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{new int[]{},new[]{1,2},new int[]{}}.SelectMany(x=>x);
__Check((flat.Count()).ToString(), "2");
