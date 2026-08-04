// vybe-test: csharp/csharp_if_else_branching/if_else_branching_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_if_else_branching.rs

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

// if_else_branching
int? maybe = 44; __P((maybe.HasValue && maybe.Value == 44).ToString());
__Check("True");
