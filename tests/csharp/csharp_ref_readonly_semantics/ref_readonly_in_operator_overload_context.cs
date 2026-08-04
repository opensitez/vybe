// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_in_operator_overload_context
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

readonly struct Num{public readonly int Value; public Num(int v){Value=v;} public static bool operator ==(Num a, ref readonly Num b)=>a.Value==b.Value; public static bool operator !=(Num a, ref readonly Num b)=>!(a==b);} var x=new Num(4); var y=new Num(4); __P((x==ref y).ToString());
__Check("True");
