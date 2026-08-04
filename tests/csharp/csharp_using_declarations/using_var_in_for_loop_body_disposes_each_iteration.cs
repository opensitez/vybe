// vybe-test: csharp/csharp_using_declarations/using_var_in_for_loop_body_disposes_each_iteration
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
for(int i=0;i<2;i++){using var x=new R("f"+i); __P(("iter").ToString());} __P(("done").ToString());
__Check("iter\nf0\niter\nf1\ndone");
