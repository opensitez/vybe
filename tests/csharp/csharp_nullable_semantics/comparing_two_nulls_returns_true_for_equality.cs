// vybe-test: csharp/csharp_nullable_semantics/comparing_two_nulls_returns_true_for_equality
// origin: languages/csharp/tests/csharp/test_csharp_nullable_semantics.rs

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

int? a=null, b=null; __P((a==b).ToString());
__Check("True");
