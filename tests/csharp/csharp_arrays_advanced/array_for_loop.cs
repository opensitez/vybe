// vybe-test: csharp/csharp_arrays_advanced/array_for_loop
// origin: languages/csharp/tests/csharp/test_csharp_arrays_advanced.rs

using static __Harness;

var arr = new[] { 1, 2, 3, 4, 5 }
;
int sum = 0;
for (int i = 0; i < arr.Length; i++) {
    sum += arr[i];
}
__P((sum).ToString());
__Check("15");

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
