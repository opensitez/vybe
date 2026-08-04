// vybe-test: csharp/csharp_using_declarations/using_var_with_continue_in_loop_still_disposes_each_time
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
foreach(var n in new[]{1,2}){using var x=new R("c"); if(n==1) continue; __P((n).ToString());} __P(("end").ToString());
__Check("c\n2\nc\nend");
