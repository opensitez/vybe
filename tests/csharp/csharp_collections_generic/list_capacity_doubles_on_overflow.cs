// vybe-test: csharp/csharp_collections_generic/list_capacity_doubles_on_overflow
// origin: languages/csharp/tests/csharp/test_csharp_collections_generic.rs

using static __Harness;

var list=new System.Collections.Generic.List<int>(4);
for(int i=0;i<8;i++) list.Add(i);
__P((list.Count).ToString());
__P((list.Capacity>=8).ToString());
__Check("8\nTrue");

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
