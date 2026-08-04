// vybe-test: csharp/csharp_oop_inheritance/method_hiding_with_new_keyword_selects_by_static_type
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

class Parent { public string Name() => "Parent"; }
class Child : Parent { public new string Name() => "Child"; }
Parent p = new Child();
__P((p.Name()).ToString());
__Check("Parent");
