// vybe-test: csharp/csharp_primary_constructors/primary_constructor_list_param_count
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((new Bag(new System.Collections.Generic.List<int> { 1, 2 }).Count).ToString());
__Check("2");

class Bag(System.Collections.Generic.List<int> items) { public int Count => items.Count; }

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
