// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_predicate_zero_with_seed
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

using static __Harness;

// `SingleOrDefault(predicate, defaultValue)`, not the reverse — the other
// order is `CS1660: Cannot convert lambda expression to type 'int'`.
__P((new[]{1,2,3}.SingleOrDefault(x=>x>10, 55)).ToString());
__Check("55");

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
