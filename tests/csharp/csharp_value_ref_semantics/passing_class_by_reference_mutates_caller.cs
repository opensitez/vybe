// vybe-test: csharp/csharp_value_ref_semantics/passing_class_by_reference_mutates_caller
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

using static __Harness;

void Mutate(C c){c.V=999;}
var c=new C{V=1}
;
Mutate(c);
__P((c.V).ToString());
__Check("999");

class C{public int V;}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
