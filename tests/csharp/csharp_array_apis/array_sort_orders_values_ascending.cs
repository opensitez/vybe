// vybe-test: csharp/csharp_array_apis/array_sort_orders_values_ascending
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var values = new[] { 4, 1, 3 }
;
System.Array.Sort(values);
foreach (var value in values) __P((value).ToString());
__Check("1\n3\n4");

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
