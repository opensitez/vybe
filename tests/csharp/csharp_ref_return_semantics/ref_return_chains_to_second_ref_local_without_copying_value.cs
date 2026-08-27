// vybe-test: csharp/csharp_ref_return_semantics/ref_return_chains_to_second_ref_local_without_copying_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_return_semantics.rs

using static __Harness;

int[] values = { 10, 20 }
;
ref int First() => ref values[0];
ref int alias = ref First();
alias = 99;
__P((values[0]).ToString());
__Check("99");

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
