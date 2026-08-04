// vybe-test: csharp/csharp_using_declarations/using_var_in_foreach_body_disposes_per_element
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

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

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__P((n).ToString());}}
foreach(var n in new[]{1,2}){using var x=new R("e"+n); __P(("each").ToString());} __P(("all").ToString());
__Check("each\ne1\neach\ne2\nall");
