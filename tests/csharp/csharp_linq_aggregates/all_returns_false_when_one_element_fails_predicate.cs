// vybe-test: csharp/csharp_linq_aggregates/all_returns_false_when_one_element_fails_predicate
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

using static __Harness;

__P((new[]{2,4,5}.All(x => x%2==0)).ToString());
__Check("False");

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
