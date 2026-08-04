// vybe-test: csharp/csharp_numeric_ops/integer_plus_double_widens_to_double
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

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

int i=3; double d=1.5;
__P((i+d).ToString());
__Check("4.5");
