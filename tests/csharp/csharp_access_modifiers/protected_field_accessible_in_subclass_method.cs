// vybe-test: csharp/csharp_access_modifiers/protected_field_accessible_in_subclass_method
// origin: languages/csharp/tests/csharp/test_csharp_access_modifiers.rs

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

class A{protected int Value=7;}
class B:A{public int Read()=>Value;}
__P((new B().Read()).ToString());
__Check("7");
