// vybe-test: csharp/csharp_generic_variance2/ienumerable_is_covariant_over_its_element_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

using static __Harness;

System.Collections.Generic.IEnumerable<string> strings=new[]{"a","b"}
;
System.Collections.Generic.IEnumerable<object> objects=strings;
__P((objects.Count()).ToString());
__Check("2");

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
