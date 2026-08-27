// vybe-test: csharp/csharp_delegates_advanced/predicate_combined_with_and
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

using static __Harness;

System.Predicate<int> positive=x=>x>0;
System.Predicate<int> even=x=>x%2==0;
System.Predicate<int> both=x=>positive(x)&&even(x);
__P((both(4)).ToString());
__P((both(-2)).ToString());
__P((both(3)).ToString());
__Check("True\nFalse\nFalse");

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
