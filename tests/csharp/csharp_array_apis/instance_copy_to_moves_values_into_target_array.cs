// vybe-test: csharp/csharp_array_apis/instance_copy_to_moves_values_into_target_array
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var source = new[] { 9, 8 }
;
var target = new int[2];
source.CopyTo(target, 0);
foreach (var value in target) __P((value).ToString());
__Check("9\n8");

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
