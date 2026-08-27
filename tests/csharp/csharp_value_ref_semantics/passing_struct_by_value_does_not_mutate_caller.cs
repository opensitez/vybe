// vybe-test: csharp/csharp_value_ref_semantics/passing_struct_by_value_does_not_mutate_caller
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

using static __Harness;

void Mutate(S s){s.V=999;}
var s=new S{V=1}
;
Mutate(s);
__P((s.V).ToString());
__Check("1");

struct S{public int V;}

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
