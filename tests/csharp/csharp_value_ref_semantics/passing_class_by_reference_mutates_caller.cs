// vybe-test: csharp/csharp_value_ref_semantics/passing_class_by_reference_mutates_caller
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

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

class C{public int V;}
void Mutate(C c){c.V=999;}
var c=new C{V=1};
Mutate(c);
__P((c.V).ToString());
__Check("999");
