// vybe-test: csharp/csharp_linq_set_ops/intersect_yields_elements_present_in_both_sequences
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

using static __Harness;

var result = new[]{1,2,3,4}
.Intersect(new[]{2,4,6}).OrderBy(x=>x);
foreach(var x in result) __P((x).ToString());
__Check("2\n4");

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
