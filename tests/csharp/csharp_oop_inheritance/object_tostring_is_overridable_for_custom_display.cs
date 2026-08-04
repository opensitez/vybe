// vybe-test: csharp/csharp_oop_inheritance/object_tostring_is_overridable_for_custom_display
// origin: languages/csharp/tests/csharp/test_csharp_oop_inheritance.rs

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

class Point { public int X,Y; public override string ToString() => $"({X},{Y})"; }
__P((new Point { X=1, Y=2 }).ToString());
__Check("(1,2)");
