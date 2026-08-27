// vybe-test: csharp/csharp_target_typed_new/target_typed_new_creates_list_without_repeating_type_arguments
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new.rs

using static __Harness;

System.Collections.Generic.List<int> values = new();
values.Add(7);
__P((values[0]).ToString());
__Check("7");

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
