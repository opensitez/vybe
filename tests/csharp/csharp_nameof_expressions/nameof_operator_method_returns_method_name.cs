// vybe-test: csharp/csharp_nameof_expressions/nameof_operator_method_returns_method_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

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

class Vector{public static Vector operator +(Vector a,Vector b)=>a; public int X;} __P((nameof(Vector.op_Addition)).ToString());
__Check("op_Addition");
