// vybe-test: csharp/csharp_nullable/null_object_pattern
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

class Box {
    public int Value;
    public Box(int v) { Value = v; }
}
Box b = null;
__P((b == null).ToString());
b = new Box(42);
__P((b == null).ToString());
__P((b.Value).ToString());
__Check("True\nFalse\n42");
