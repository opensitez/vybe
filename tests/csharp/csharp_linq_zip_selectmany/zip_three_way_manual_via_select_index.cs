// vybe-test: csharp/csharp_linq_zip_selectmany/zip_three_way_manual_via_select_index
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a=new[]{1,2,3}; var b=new[]{4,5,6};
var z=a.Zip(b,(x,y)=>x+y);
__Check((z.ElementAt(1)).ToString(), "7");
