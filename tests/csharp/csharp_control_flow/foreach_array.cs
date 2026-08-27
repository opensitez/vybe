// vybe-test: csharp/csharp_control_flow/foreach_array
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

using static __Harness;

int[] arr = {10, 20, 30}
;
int sum = 0;
foreach (var x in arr) {
    sum += x;
}
__P((sum).ToString());
__Check("60");

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
