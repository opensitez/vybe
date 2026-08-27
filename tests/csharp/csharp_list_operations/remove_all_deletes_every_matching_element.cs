// vybe-test: csharp/csharp_list_operations/remove_all_deletes_every_matching_element
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

using static __Harness;

var list = new System.Collections.Generic.List<int>{1,2,3,4,5}
;
list.RemoveAll(x => x % 2 == 0);
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
