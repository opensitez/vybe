// vybe-test: csharp/csharp_linq_zip_selectmany/zip_both_empty_yields_zero_pairs
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=System.Array.Empty<int>().Zip(System.Array.Empty<int>(),(a,b)=>a+b);
__Check((z.Count()).ToString(), "0");
