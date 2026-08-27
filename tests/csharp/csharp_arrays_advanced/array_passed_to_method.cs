// vybe-test: csharp/csharp_arrays_advanced/array_passed_to_method
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

using static __Harness;

var nums = new[] { 1, 2, 3, 4 }
;
__P((Utils.Sum(nums)).ToString());
__Check("10");

class Utils {
    public static int Sum(int[] arr) {
        int total = 0;
        foreach (var x in arr) total += x;
        return total;
    }
}

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
