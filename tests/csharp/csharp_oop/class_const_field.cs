// vybe-test: csharp/csharp_oop/class_const_field
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

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

class MathConst {
    public const double PI = 3.14159;
    public const double E = 2.71828;
}
__P((MathConst.PI).ToString());
__P((MathConst.E).ToString());
__Check("3.14159\n2.71828");
