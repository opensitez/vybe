// vybe-test: csharp/csharp_reflection/property_info_get_value_reads_property_dynamically
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

var item = new Item { Id=7 }
;
var prop = typeof(Item).GetProperty("Id");
__P((prop.GetValue(item)).ToString());
__Check("7");

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
