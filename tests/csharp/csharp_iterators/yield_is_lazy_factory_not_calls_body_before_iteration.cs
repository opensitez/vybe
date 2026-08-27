// vybe-test: csharp/csharp_iterators/yield_is_lazy_factory_not_calls_body_before_iteration
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

using static __Harness;

int calls=0;
System.Collections.Generic.IEnumerable<int> Lazy() {
    calls++;
    yield return 1;
}
__P((calls).ToString());
var seq = Lazy();
__P((calls).ToString());
foreach(var _ in seq) {}
__P((calls).ToString());
__Check("0\n0\n1");

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
