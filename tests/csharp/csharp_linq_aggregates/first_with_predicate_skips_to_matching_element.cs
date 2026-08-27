// vybe-test: csharp/csharp_linq_aggregates/first_with_predicate_skips_to_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

using static __Harness;

__P((new[]{1,2,3,4}.First(x => x>2)).ToString());
__Check("3");

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
