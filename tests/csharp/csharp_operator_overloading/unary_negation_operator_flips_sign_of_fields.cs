// vybe-test: csharp/csharp_operator_overloading/unary_negation_operator_flips_sign_of_fields
// origin: languages/csharp/tests/csharp/test_csharp_operator_overloading.rs

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

struct Vec{public int X;
public static Vec operator-(Vec v)=>new Vec{X=-v.X};}
var v=-new Vec{X=7};
__P((v.X).ToString());
__Check("-7");
