// vybe-test: csharp/csharp_local_function_static/local_function_while_loop_with_capture
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

int Count(int n){int i=0; int acc=0; while(i<n){int Step(int x)=>acc+x; acc=Step(i+1); i++;} return acc;} __P((Count(3)).ToString());
__Check("6");
