// vybe-test: csharp/csharp_linq_projections/to_list_materializes_query_to_mutable_list
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var list = new[]{1,2,3}
.Select(x => x*2).ToList();
__P((list.GetType().Name).ToString());
__Check("List`1");

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
