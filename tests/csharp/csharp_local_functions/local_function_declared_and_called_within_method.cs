// vybe-test: csharp/csharp_local_functions/local_function_declared_and_called_within_method
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

int Square(int n){
    int Sq(int x)=>x*x;
    return Sq(n);
}
__P((Square(5)).ToString());
__Check("25");
