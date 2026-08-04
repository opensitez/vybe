// vybe-test: csharp/csharp_nullable/null_conditional_member
// origin: languages/csharp/tests/csharp/test_csharp_nullable.rs

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

class Wrapper {
    public string Value;
    public Wrapper(string v) { Value = v; }
}
Wrapper w = null;
__P((w?.Value ?? "null").ToString());
w = new Wrapper("hello");
__P((w?.Value ?? "null").ToString());
__Check("null\nhello");
