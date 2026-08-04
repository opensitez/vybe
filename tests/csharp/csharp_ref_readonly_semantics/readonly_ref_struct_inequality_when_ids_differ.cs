// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_inequality_when_ids_differ
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

readonly ref struct Tag{public readonly int Id; public Tag(int id){Id=id;} public bool Equals(Tag other)=>Id==other.Id;} var a=new Tag(1); var b=new Tag(2); __P((a.Equals(b)).ToString());
__Check("False");
