// vybe-test: csharp/csharp_linq_numeric/long_count_works_on_large_range
// origin: languages/csharp/tests/csharp/test_csharp_linq_numeric.rs

using static __Harness;

long c=Enumerable.Range(0,1000).LongCount();
__P((c).ToString());
__Check("1000");

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
