// vybe-test: csharp/csharp_local_function_static/static_local_function_clamp_high
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

int Clamp(int n,int max){static int Cap(int x,int m)=>x>m?m:x; return Cap(n,max);} __P((Clamp(15,10)).ToString());
__Check("10");
