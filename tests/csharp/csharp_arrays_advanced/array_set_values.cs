// vybe-test: csharp/csharp_arrays_advanced/array_set_values
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

using static __Harness;

var arr = new int[3];
arr[0] = 10;
arr[1] = 20;
arr[2] = 30;
__P((arr[0] + arr[1] + arr[2]).ToString());
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
