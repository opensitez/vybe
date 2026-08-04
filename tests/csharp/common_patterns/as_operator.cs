// vybe-test: csharp/common_patterns/as_operator
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

object obj = "hello";
string s = obj as string;
__P((s != null ? s : "null").ToString());
int? i = obj as int?;
__P((i != null ? i.ToString() : "null").ToString());
__Check("hello\nnull");
