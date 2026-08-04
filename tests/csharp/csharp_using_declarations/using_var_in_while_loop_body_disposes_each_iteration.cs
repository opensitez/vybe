// vybe-test: csharp/csharp_using_declarations/using_var_in_while_loop_body_disposes_each_iteration
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
int i=0; while(i<2){using var x=new R(i.ToString()); __P(("loop").ToString()); i++;} __P(("exit").ToString());
__Check("loop\n0\nloop\n1\nexit");
