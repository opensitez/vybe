// vybe-test: csharp/more_classes/auto_property_set_after_construction
// origin: languages/csharp/tests/csharp/test_more_classes.rs

using static __Harness;

var item = new Item();
item.Name = "Widget";
__P((item.Name).ToString());
__Check("Widget");

class Item {
            public string Name { get; set; }
            public Item() {}
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
