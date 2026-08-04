// vybe-test: csharp/csharp_oop_inheritance/protected_member_is_accessible_in_derived_class
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

class Base { protected int Secret = 42; }
class Child : Base { public int Get() => Secret; }
__P((new Child().Get()).ToString());
__Check("42");
