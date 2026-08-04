// vybe-test: csharp/csharp_oop_inheritance/base_call_delegates_to_parent_implementation
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

class A { public virtual string Greet() => "Hello"; }
class B : A { public override string Greet() => base.Greet() + " World"; }
__P((new B().Greet()).ToString());
__Check("Hello World");
