// vybe-test: csharp/csharp_local_function_static/local_function_in_loop_accumulates
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

int Sum(int n){int total=0; for(int i=1;i<=n;i++){int Add(int x)=>total+x; total=Add(i);} return total;} __P((Sum(3)).ToString());
__Check("6");
