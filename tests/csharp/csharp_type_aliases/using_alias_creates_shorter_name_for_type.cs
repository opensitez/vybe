// vybe-test: csharp/csharp_type_aliases/using_alias_creates_shorter_name_for_type
// origin: languages/csharp/tests/csharp/test_csharp_type_aliases.rs

using static __Harness;
using IntList=System.Collections.Generic.List<int>;

var list=new IntList{1,2,3}
;
__P((list.Count).ToString());
__Check("3");

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
