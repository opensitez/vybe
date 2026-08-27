// vybe-test: csharp/csharp_object_initializers/collection_initializer_populates_list_inline
// origin: languages/csharp/tests/csharp/test_csharp_object_initializers.rs

using static __Harness;

var list=new System.Collections.Generic.List<int>{10,20,30}
;
__P((list[1]).ToString());
__Check("20");

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
