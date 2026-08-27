// vybe-test: csharp/csharp_new_features/target_typed_new_in_constructor_argument
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

using static __Harness;

var b = new Box(new());
b.Items.Add(9);
__P((b.Items.Count).ToString());
__Check("1");

class Box { public System.Collections.Generic.List<int> Items; public Box(System.Collections.Generic.List<int> i){Items=i;} }

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
