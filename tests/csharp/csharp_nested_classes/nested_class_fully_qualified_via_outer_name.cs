// vybe-test: csharp/csharp_nested_classes/nested_class_fully_qualified_via_outer_name
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

using static __Harness;

var item=new Container.Item();
__P((item.Value).ToString());
__Check("7");

class Container{public class Item{public int Value=7;}}

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
