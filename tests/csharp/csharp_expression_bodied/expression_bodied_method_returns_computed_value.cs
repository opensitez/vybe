// vybe-test: csharp/csharp_expression_bodied/expression_bodied_method_returns_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

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

class Calc{public int Double(int n)=>n*2;}
__P((new Calc().Double(5)).ToString());
__Check("10");
