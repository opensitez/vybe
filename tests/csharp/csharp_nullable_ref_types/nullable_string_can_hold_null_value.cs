// vybe-test: csharp/csharp_nullable_ref_types/nullable_string_can_hold_null_value
// origin: languages/csharp/tests/csharp/test_csharp_nullable_ref_types.rs

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

string? s=null;
__P((s==null).ToString());
__Check("True");
