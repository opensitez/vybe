// vybe-test: csharp/csharp_structs_value_semantics/passing_struct_by_value_does_not_mutate_caller_copy
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

void Bump(Counter counter) { counter.Value++; }
var counter = new Counter { Value = 2 }
;
Bump(counter);
__P((counter.Value).ToString());
__Check("2");

struct Counter { public int Value; }

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
