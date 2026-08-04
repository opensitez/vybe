// vybe-test: csharp/csharp_linq_zip_selectmany/select_many_three_nested_levels_count
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

var data=new[]{new[]{new[]{1,2}},new[]{new[]{3}}};
var flat=data.SelectMany(a=>a).SelectMany(b=>b);
__P((flat.Count()).ToString());
__Check("3");
