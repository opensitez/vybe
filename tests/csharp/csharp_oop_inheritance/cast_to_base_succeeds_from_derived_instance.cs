// vybe-test: csharp/csharp_oop_inheritance/cast_to_base_succeeds_from_derived_instance
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

class Base { public int X = 1; }
class Derived : Base { public int Y = 2; }
Base b = (Base)new Derived();
__P((b.X).ToString());
__Check("1");
