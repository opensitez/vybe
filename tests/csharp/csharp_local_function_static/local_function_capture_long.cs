// vybe-test: csharp/csharp_local_function_static/local_function_capture_long
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

long baseVal=10000000000L; int Add(int n){int A(int x)=>x+(int)(baseVal%100); return A(n);} __P((Add(5)).ToString());
__Check("5");
