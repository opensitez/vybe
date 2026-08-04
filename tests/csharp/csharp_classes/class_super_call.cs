// vybe-test: csharp/csharp_classes/class_super_call
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

class Base {
    public virtual string Greet() { return "Hello"; }
}
class Derived : Base {
    public override string Greet() { return base.Greet() + " World"; }
}
var d = new Derived();
__P((d.Greet()).ToString());
__Check("Hello World");
