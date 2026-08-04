// vybe-test: csharp/csharp_enum_operations/enum_explicit_value_tostring_probe
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

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

enum Num{X=7} __P((Num.X.ToString()).ToString());
__Check("X");
