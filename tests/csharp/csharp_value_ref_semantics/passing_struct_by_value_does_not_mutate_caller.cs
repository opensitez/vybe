// vybe-test: csharp/csharp_value_ref_semantics/passing_struct_by_value_does_not_mutate_caller
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

struct S{public int V;}
void Mutate(S s){s.V=999;}
var s=new S{V=1};
Mutate(s);
__P((s.V).ToString());
__Check("1");
