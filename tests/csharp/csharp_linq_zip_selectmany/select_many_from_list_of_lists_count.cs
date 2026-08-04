// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_from_list_of_lists_count
// origin: languages/csharp/tests/csharp/test_csharp_linq_zip_selectmany.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var lists=new System.Collections.Generic.List<int[]>{
    new[]{1,2},new[]{3}}; 
__P((lists.SelectMany(x=>x).Count()).ToString());
__Check("3");
