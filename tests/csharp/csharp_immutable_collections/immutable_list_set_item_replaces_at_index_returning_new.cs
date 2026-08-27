// vybe-test: csharp/csharp_immutable_collections/immutable_list_set_item_replaces_at_index_returning_new
// origin: languages/csharp/tests/csharp/test_csharp_immutable_collections.rs

using static __Harness;

var list=System.Collections.Immutable.ImmutableList.Create(1,2,3);
var updated=list.SetItem(1,99);
__P((list[1]).ToString());
__P((updated[1]).ToString());
__Check("2\n99");

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
