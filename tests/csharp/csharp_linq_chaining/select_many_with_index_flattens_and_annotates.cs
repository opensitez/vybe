// vybe-test: csharp/csharp_linq_chaining/select_many_with_index_flattens_and_annotates
// origin: languages/csharp/tests/csharp/test_csharp_linq_chaining.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var groups=new[]{new[]{1,2},new[]{3,4}};
var result=groups.SelectMany((g,i)=>g.Select(x=>i*10+x));
__Check((string.Join(",",result)).ToString(), "1,2,13,14");
