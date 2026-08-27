// vybe-test: csharp/csharp_reflection/property_info_set_value_writes_property_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

var item = new Item();
var prop = typeof(Item).GetProperty("Id");
prop.SetValue(item, 99);
__P((item.Id).ToString());
__Check("99");

class Item { public int Id {get;set;} }

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
