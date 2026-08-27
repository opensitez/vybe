// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_of_custom_type_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Collections.Generic.List<Item> items = new();
items.Add(new Item { Name = "tool" });
__P((items[0].Name).ToString());
__Check("tool");

class Item { public string Name = ""; }

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
