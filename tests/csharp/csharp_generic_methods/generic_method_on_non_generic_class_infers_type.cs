// vybe-test: csharp/csharp_generic_methods/generic_method_on_non_generic_class_infers_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_methods.rs

using static __Harness;

__P((Utils.First(new[]{10,20,30})).ToString());
__P((Utils.First(new[]{"a","b"})).ToString());
__Check("10\na");

class Utils{public static T First<T>(T[] arr)=>arr[0];}

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
