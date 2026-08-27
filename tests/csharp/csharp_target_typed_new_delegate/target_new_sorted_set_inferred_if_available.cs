// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_sorted_set_inferred_if_available
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Collections.Generic.SortedSet<int> ordered = new();
ordered.Add(3);
ordered.Add(1);
ordered.Add(2);
foreach (var n in ordered) __P((n).ToString());
__Check("1\n2\n3");

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
