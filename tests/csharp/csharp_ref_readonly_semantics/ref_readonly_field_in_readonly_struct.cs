// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_field_in_readonly_struct
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

readonly struct Pair{public readonly int First; public readonly int Second; public Pair(int a,int b){First=a; Second=b;}} var p=new Pair(2,3); __P((p.First+p.Second).ToString());
__Check("5");
