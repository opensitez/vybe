// vybe-test: csharp/csharp_linq_projections/take_returns_first_n_elements
// origin: languages/csharp/tests/csharp/test_csharp_linq_projections.rs

using static __Harness;

var result = new[]{10,20,30,40}
.Take(2);
foreach(var n in result) __P((n).ToString());
__Check("10\n20");

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
