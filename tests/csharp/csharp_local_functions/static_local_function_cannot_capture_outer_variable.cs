// vybe-test: csharp/csharp_local_functions/static_local_function_cannot_capture_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

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

static int Pure(int a,int b){
    static int Add(int x,int y)=>x+y;
    return Add(a,b);
}
__P((Pure(4,5)).ToString());
__Check("9");
