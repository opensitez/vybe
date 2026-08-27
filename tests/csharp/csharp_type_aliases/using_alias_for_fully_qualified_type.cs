// vybe-test: csharp/csharp_type_aliases/using_alias_for_fully_qualified_type
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

using static __Harness;
using Dict=System.Collections.Generic.Dictionary<string,int>;

var d=new Dict{{"a",1},{"b",2}}
;
__P((d["b"]).ToString());
__Check("2");

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
