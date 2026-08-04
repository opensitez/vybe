// vybe-test: csharp/csharp_value_ref_semantics/boxing_wraps_value_type_in_object
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

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

int n=42; object o=n;
__P((o).ToString()); __P((o is int).ToString());
__Check("42\nTrue");
