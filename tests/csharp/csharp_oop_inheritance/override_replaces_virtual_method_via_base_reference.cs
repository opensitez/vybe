// vybe-test: csharp/csharp_oop_inheritance/override_replaces_virtual_method_via_base_reference
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

class Animal { public virtual string Sound() => "..."; }
class Dog : Animal { public override string Sound() => "woof"; }
Animal a = new Dog();
__P((a.Sound()).ToString());
__Check("woof");
