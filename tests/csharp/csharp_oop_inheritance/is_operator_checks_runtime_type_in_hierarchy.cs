// vybe-test: csharp/csharp_oop_inheritance/is_operator_checks_runtime_type_in_hierarchy
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

class A { }
class B : A { }
object obj = new B();
__P((obj is A).ToString());
__P((obj is B).ToString());
__Check("True\nTrue");
