// vybe-test: csharp/csharp_using_declarations/using_var_chained_scope_three_levels_lifo
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
{using var l1=new R("l1"); {using var l2=new R("l2"); {using var l3=new R("l3"); __P(("3").ToString());}} __P(("2").ToString());} __P(("1").ToString());
__Check("3\nl3\n2\nl2\n1\nl1");
