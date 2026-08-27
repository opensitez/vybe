// vybe-test: csharp/csharp_linq_set_ops/zip_pairs_elements_from_two_sequences_by_position
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

using static __Harness;

var result = new[]{1,2,3}
.Zip(new[]{10,20,30}, (a,b) => a*b);
foreach(var x in result) __P((x).ToString());
__Check("10\n40\n90");

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
