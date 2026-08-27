// vybe-test: csharp/csharp_array_apis/array_convert_all_maps_values_to_new_type
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

using static __Harness;

var values = new[] { 1, 2, 3 }
;
var text = System.Array.ConvertAll(values, value => "n" + value);
foreach (var value in text) __P((value).ToString());
__Check("n1\nn2\nn3");

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
