// vybe-test: csharp/csharp_exceptions_flow/finally_always_runs_even_after_return
// origin: languages/csharp/tests/csharp/test_csharp_exceptions_flow.rs

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

bool ran=false;
int Compute(){
    try{return 42;}
    finally{ran=true;}
}
int v=Compute();
__P((v).ToString()); __P((ran).ToString());
__Check("42\nTrue");
