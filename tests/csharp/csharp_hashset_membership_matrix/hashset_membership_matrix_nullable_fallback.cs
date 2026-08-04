// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

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

// hashset_membership_matrix
int? maybe = null; int fallback = maybe ?? 33; __P((fallback == 33).ToString());
__Check("True");
