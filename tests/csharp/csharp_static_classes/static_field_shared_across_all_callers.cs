// vybe-test: csharp/csharp_static_classes/static_field_shared_across_all_callers
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

using static __Harness;

Counter.Count++;
Counter.Count++;
__P((Counter.Count).ToString());
__Check("2");

static class Counter { public static int Count = 0; }

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
