// vybe-test: csharp/csharp_local_function_static/static_local_function_min_of_two
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

int Min(int a,int b){static int Pick(int x,int y)=>x<y?x:y; return Pick(a,b);} __P((Min(3,9)).ToString());
__Check("3");
