// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_from_list_of_lists_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var lists=new System.Collections.Generic.List<int[]>{
    new[]{1,2},new[]{3}}; 
__Check((lists.SelectMany(x=>x).Count()).ToString(), "3");
