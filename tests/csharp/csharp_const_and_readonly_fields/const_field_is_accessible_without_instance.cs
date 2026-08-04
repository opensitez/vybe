// vybe-test: csharp/csharp_const_and_readonly_fields/const_field_is_accessible_without_instance
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

class Limits {
    public const int Max = 100;
}
__P((Limits.Max).ToString());
__Check("100");
