// vybe-test: csharp/csharp_delegates_advanced/func_returns_func_partial_application
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

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

System.Func<int,System.Func<int,int>> multiply=factor=>n=>n*factor;
var triple=multiply(3);
__P((triple(7)).ToString());
__Check("21");
