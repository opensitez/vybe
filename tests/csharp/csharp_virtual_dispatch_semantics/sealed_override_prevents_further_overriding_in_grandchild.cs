// vybe-test: csharp/csharp_virtual_dispatch_semantics/sealed_override_prevents_further_overriding_in_grandchild
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

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
    public virtual string Tag() { return "base"; }
}
class Middle : Base {
    public sealed override string Tag() { return "middle"; }
}
class Leaf : Middle {
    public override string Tag() { return "leaf"; }
}
Base item = new Leaf();
__P((item.Tag()).ToString());
__Check("middle");
