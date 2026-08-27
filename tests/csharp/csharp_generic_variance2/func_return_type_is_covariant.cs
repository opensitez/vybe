// vybe-test: csharp/csharp_generic_variance2/func_return_type_is_covariant
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

using static __Harness;

System.Func<string> getStr=()=>"hello";
System.Func<object> getObj=getStr;
__P((getObj()).ToString());
__Check("hello");

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
