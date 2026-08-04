// vybe-test: csharp/csharp_using_declarations/using_var_in_do_while_executes_at_least_once_then_disposes
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
int n=0; do{using var x=new R("do"); __P(("once").ToString()); n++;} while(n<1); __P(("fin").ToString());
__Check("once\ndo\nfin");
