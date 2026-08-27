// vybe-test: csharp/csharp_list_operations/add_range_appends_all_elements_of_collection
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

using static __Harness;

var list = new System.Collections.Generic.List<int>{1}
;
list.AddRange(new[]{2,3,4});
__P((list.Count).ToString());
__Check("4");

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
