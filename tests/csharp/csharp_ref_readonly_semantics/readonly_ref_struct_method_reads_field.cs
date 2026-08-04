// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_method_reads_field
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

readonly ref struct Counter{public readonly int Value; public Counter(int v){Value=v;} public int Doubled()=>Value*2;} var c=new Counter(6); __P((c.Doubled()).ToString());
__Check("12");
