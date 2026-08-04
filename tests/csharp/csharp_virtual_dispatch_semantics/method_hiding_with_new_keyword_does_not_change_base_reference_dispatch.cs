// vybe-test: csharp/csharp_virtual_dispatch_semantics/method_hiding_with_new_keyword_does_not_change_base_reference_dispatch
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
    public string Name() { return "base"; }
}
class Derived : Base {
    public new string Name() { return "derived"; }
}
Base reference = new Derived();
Derived concrete = new Derived();
__P((reference.Name()).ToString());
__P((concrete.Name()).ToString());
__Check("base\nderived");
