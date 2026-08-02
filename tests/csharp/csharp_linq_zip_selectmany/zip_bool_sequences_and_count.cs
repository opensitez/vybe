// vybe-test: csharp/csharp_linq_zip_selectmany/zip_bool_sequences_and_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var z=new[]{true,false,true}.Zip(new[]{false,true,false},(a,b)=>a&&b);
__Check((z.Count(x=>x)).ToString(), "0");
