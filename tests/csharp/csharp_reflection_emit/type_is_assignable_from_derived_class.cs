// vybe-test: csharp/csharp_reflection_emit/type_is_assignable_from_derived_class
// origin: languages/csharp/tests/csharp/test_csharp_reflection_emit.rs

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

class A{} class B:A{}
__P((typeof(A).IsAssignableFrom(typeof(B))).ToString());
__P((typeof(B).IsAssignableFrom(typeof(A))).ToString());
__Check("True\nFalse");
