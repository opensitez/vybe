// vybe-test: csharp/csharp_using_declarations/using_var_disposal_order_with_interleaved_console_writes
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

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__P(("d:"+n).ToString());}}
using var a=new R("a"); __P(("m1").ToString()); using var b=new R("b"); __P(("m2").ToString());
__Check("m1\nm2\nd:b\nd:a");
