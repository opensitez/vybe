// vybe-test: csharp/csharp_static_constructor_once/static_constructor_increments_counter_only_once_across_two_instance_allocations
// origin: languages/csharp/tests/csharp/test_csharp_static_constructor_once.rs

using static __Harness;

_ = new Tracker();
_ = new Tracker();
__P((Tracker.Instances).ToString());
__Check("1");

class Tracker {
    public static int Instances;
    static Tracker() { Instances++; }
}

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
