// vybe-test: csharp/csharp_object_equality/null_equals_null_returns_true_via_static_method
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

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

__P((object.Equals(null, null)).ToString());
__Check("True");
