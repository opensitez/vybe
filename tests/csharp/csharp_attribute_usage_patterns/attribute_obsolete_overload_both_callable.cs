// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_overload_both_callable
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

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

using System; class S{[Obsolete("a")] public int Go()=>1; [Obsolete("b")] public int Go(int x)=>x;} __P((new S().Go()).ToString()); __P((new S().Go(5)).ToString());
__Check("1\n5");
