// vybe-test: csharp/csharp_local_function_static/local_function_nested_static_inner
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

int Calc(int n){static int Inner(int x)=>x+5; int Outer(int v)=>Inner(v)*2; return Outer(n);} __P((Calc(3)).ToString());
__Check("16");
