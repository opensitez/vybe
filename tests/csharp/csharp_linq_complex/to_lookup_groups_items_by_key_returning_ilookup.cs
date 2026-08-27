// vybe-test: csharp/csharp_linq_complex/to_lookup_groups_items_by_key_returning_ilookup
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

using static __Harness;

var data=new[]{(K:"a",V:1),(K:"a",V:2),(K:"b",V:3)}
;
var lu=data.ToLookup(x=>x.K,x=>x.V);
__P((lu["a"].Sum()).ToString());
__P((lu["b"].Sum()).ToString());
__Check("3\n3");

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
