// vybe-test: csharp/csharp_struct_features/struct_passed_by_ref_mutates_caller_copy
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

using static __Harness;

void Increment(ref Counter c) { c.N++; }
var c = new Counter { N=5 }
;
Increment(ref c);
__P((c.N).ToString());
__Check("6");

struct Counter { public int N; }

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
