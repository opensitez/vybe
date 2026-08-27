// vybe-test: csharp/csharp_constructor_chains/field_initializer_runs_before_constructor_body_reads_value
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

using static __Harness;

new Box();
__Check("init");

class Box { string name = "init"; public Box() { __P((name).ToString()); } }

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
