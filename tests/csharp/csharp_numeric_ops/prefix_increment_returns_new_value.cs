// vybe-test: csharp/csharp_numeric_ops/prefix_increment_returns_new_value
// origin: languages/csharp/tests/csharp/test_csharp_numeric_ops.rs

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

int x=5; int y=++x;
__P((y).ToString()); __P((x).ToString());
__Check("6\n6");
