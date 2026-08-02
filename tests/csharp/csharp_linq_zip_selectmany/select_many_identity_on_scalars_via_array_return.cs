// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_identity_on_scalars_via_array_return
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var flat=new[]{1,2,3}.SelectMany(n=>new[]{n,n});
__Check((flat.Count()).ToString(), "6");
