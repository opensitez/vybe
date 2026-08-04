// vybe-test: csharp/csharp_classes/multi_level_inheritance
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

class A {
    public virtual string Who() { return "A"; }
}
class B : A {
    public override string Who() { return "B->" + base.Who(); }
}
class C : B {
    public override string Who() { return "C->" + base.Who(); }
}
var c = new C();
__P((c.Who()).ToString());
__Check("C->B->A");
