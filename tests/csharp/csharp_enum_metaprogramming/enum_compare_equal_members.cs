// vybe-test: csharp/csharp_enum_metaprogramming/enum_compare_equal_members
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

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

enum Eq{X,Y} __P((Eq.X==Eq.X).ToString()); __P((Eq.X==Eq.Y).ToString());
__Check("True\nFalse");
