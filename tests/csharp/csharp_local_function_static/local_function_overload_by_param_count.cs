// vybe-test: csharp/csharp_local_function_static/local_function_overload_by_param_count
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

int Compute(int n){int One(int x)=>x+1; int Two(int x,int y)=>x+y; return Two(n,One(n));} __P((Compute(5)).ToString());
__Check("11");
