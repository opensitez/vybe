// vybe-test: csharp/csharp_oop_inheritance/abstract_class_provides_partial_implementation
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

abstract class Base {
    public abstract int Value();
    public int Double() => Value() * 2;
}
class Impl : Base { public override int Value() => 5; }
__P((new Impl().Double()).ToString());
__Check("10");
