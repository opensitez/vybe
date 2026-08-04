// vybe-test: csharp/csharp_classes/class_pass_by_reference
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

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
}
void Modify(Box b) {
    b.Value = 99;
}
var b = new Box();
b.Value = 1;
Modify(b);
__P((b.Value).ToString());
__Check("99");
