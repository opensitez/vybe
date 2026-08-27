// vybe-test: csharp/csharp_deconstruction_patterns/deconstruction_in_foreach_loop_over_tuple_array
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

using static __Harness;

var pairs = new[]{(1,"a"),(2,"b"),(3,"c")}
;
int sum=0;
foreach(var (n, _) in pairs) sum+=n;
__P((sum).ToString());
__Check("6");

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
