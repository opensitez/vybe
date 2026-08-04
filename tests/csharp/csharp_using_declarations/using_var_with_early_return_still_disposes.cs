// vybe-test: csharp/csharp_using_declarations/using_var_with_early_return_still_disposes
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
int Go(bool stop){using var x=new R("go"); if(stop) return 1; return 2;} __P((Go(true)).ToString());
__Check("go\n1");
