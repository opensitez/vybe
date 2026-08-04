// vybe-test: csharp/csharp_const_and_readonly_fields/static_readonly_vs_const_both_accessible_on_type_name
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

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

class Mix {
    public const int A = 1;
    public static readonly int B = 2;
}
__P((Mix.A + Mix.B).ToString());
__Check("3");
