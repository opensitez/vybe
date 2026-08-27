// vybe-test: csharp/csharp_class_features/static_field_shared_across_all_instances
// origin: languages/csharp/tests/csharp/test_csharp_class_features.rs

using static __Harness;

new Ctr();
new Ctr();
new Ctr();
__P((Ctr.Count).ToString());
__Check("3");

class Ctr{public static int Count=0; public Ctr(){Count++;}}

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
