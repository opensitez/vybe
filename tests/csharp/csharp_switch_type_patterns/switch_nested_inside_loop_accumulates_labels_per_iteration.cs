// vybe-test: csharp/csharp_switch_type_patterns/switch_nested_inside_loop_accumulates_labels_per_iteration
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

using static __Harness;

string trace = "";
for (int i = 0; i < 3; i++) {
    trace += i switch { 0 => "a", 1 => "b", _ => "c" };
}
__P((trace).ToString());
__Check("abc");

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
