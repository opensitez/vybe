// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_of_struct_field
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

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

struct Widget{public int Id;} Widget w=new Widget(); w.Id=7; ref readonly int Get(ref Widget item)=>ref item.Id; __P((Get(ref w)).ToString());
__Check("7");
