// vybe-test: csharp/csharp_array_apis/array_copy_moves_values_between_arrays
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var source = new[] { 5, 6, 7 }
;
var target = new int[3];
System.Array.Copy(source, target, 3);
foreach (var value in target) __P((value).ToString());
__Check("5\n6\n7");

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
