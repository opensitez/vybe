// vybe-test: csharp/csharp_nameof_expressions/nameof_partial_class_member_returns_member_name
// origin: languages/csharp/tests/csharp/test_csharp_nameof_expressions.rs

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

partial class Partial{public int Id;} __P((nameof(Partial.Id)).ToString());
__Check("Id");
