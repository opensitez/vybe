// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_obsolete_and_flags_on_different_types
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

using System; [Flags] enum F{A=1} [Obsolete("old")] class S{public int Use()=>(int)F.A;} __P((new S().Use()).ToString());
__Check("1");
