// vybe-test: csharp/csharp_threading_primitives/thread_static_field_defaults_to_zero_on_main_thread
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

using static __Harness;

__P((Counter.Value).ToString());
__Check("0");

class Counter {
    [System.ThreadStatic]
    public static int Value;
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
