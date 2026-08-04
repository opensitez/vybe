// vybe-test: csharp/csharp_reflection_emit/type_base_type_reflects_inheritance_chain
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

class A{} class B:A{} class C:B{}
__P((typeof(C).BaseType.Name).ToString());
__P((typeof(C).BaseType.BaseType.Name).ToString());
__Check("B\nA");
