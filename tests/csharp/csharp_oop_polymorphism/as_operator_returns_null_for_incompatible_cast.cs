// vybe-test: csharp/csharp_oop_polymorphism/as_operator_returns_null_for_incompatible_cast
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

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

class A{} class B{}
object o=new A();
__P((o as B==null).ToString());
__Check("True");
