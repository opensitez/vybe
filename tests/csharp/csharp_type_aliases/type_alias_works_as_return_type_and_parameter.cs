// vybe-test: csharp/csharp_type_aliases/type_alias_works_as_return_type_and_parameter
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

using static __Harness;
using NameMap=System.Collections.Generic.Dictionary<string,string>;

NameMap Build()=>new NameMap{{"k","v"}}
;
__P((Build()["k"]).ToString());
__Check("v");

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
