// vybe-test: csharp/csharp_local_functions/local_function_captures_outer_variable
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

int multiplier=3;
int Mul(int n){
    int Scaled(int x)=>x*multiplier;
    return Scaled(n);
}
__P((Mul(7)).ToString());
__Check("21");
