// vybe-test: csharp/csharp_covariance_contravariance/func_return_type_covariance_allows_derived_func_in_base_func
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

using static __Harness;

System.Func<string> getString = () => "hi";
System.Func<object> getObject = getString;
__P((getObject()).ToString());
__Check("hi");

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
