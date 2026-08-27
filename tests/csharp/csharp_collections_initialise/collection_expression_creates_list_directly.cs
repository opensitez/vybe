// vybe-test: csharp/csharp_collections_initialise/collection_expression_creates_list_directly
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

using static __Harness;

System.Collections.Generic.List<int> list=[1,2,3];
__P((list.Count).ToString());
__P((list[1]).ToString());
__Check("3\n2");

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
