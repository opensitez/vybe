// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_can_be_local_variable
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

readonly ref struct Marker{public readonly int Code; public Marker(int c){Code=c;}} var m=new Marker(42); __P((m.Code).ToString());
__Check("42");
