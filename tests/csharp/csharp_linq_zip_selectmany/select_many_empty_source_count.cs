// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_empty_source_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=System.Array.Empty<int[]>().SelectMany(x=>x);
__Check((flat.Count()).ToString(), "0");
