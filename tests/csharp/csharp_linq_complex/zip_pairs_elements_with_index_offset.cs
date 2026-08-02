// vybe-test: csharp/csharp_linq_complex/zip_pairs_elements_with_index_offset
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new[]{1,2,3}; var b=new[]{4,5,6};
var r=a.Zip(b,(x,y)=>x*y);
__Check((r.Sum()).ToString(), "32");
