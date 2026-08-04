// vybe-test: csharp/oop_advanced/virtual_override_three_levels
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

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
    public override string Who() { return "B"; }
}
class C : B {
    public override string Who() { return "C"; }
}
A obj = new C();
__P((obj.Who()).ToString());
__Check("C");
