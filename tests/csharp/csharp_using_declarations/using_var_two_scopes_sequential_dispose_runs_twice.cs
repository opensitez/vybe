// vybe-test: csharp/csharp_using_declarations/using_var_two_scopes_sequential_dispose_runs_twice
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

class R:System.IDisposable{public static int N=0;public void Dispose(){N++;}}
{using var x=new R();} {using var y=new R();} __P((R.N).ToString());
__Check("2");
