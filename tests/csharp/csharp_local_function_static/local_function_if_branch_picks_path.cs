// vybe-test: csharp/csharp_local_function_static/local_function_if_branch_picks_path
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

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

string Sign(int n){string Pos(int x)=>"+"; string Neg(int x)=>"-"; if(n>=0){return Pos(n);} return Neg(n);} __P((Sign(-1)).ToString());
__Check("-");
