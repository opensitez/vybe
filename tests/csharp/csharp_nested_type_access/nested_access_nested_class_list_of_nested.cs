// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_class_list_of_nested
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

using static __Harness;

__P((new Bag().All()[0].Id).ToString());
__Check("1");

class Bag{public class Item{public int Id;} public System.Collections.Generic.List<Item> All(){var list=new System.Collections.Generic.List<Item>(); list.Add(new Item{Id=1}); return list;}}

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
