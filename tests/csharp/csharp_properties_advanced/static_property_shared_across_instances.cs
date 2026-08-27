// vybe-test: csharp/csharp_properties_advanced/static_property_shared_across_instances
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

using static __Harness;

AppConfig.Version="2.0";
__P((new System.Object().GetType()!=null).ToString());
__P((AppConfig.Version).ToString());
__Check("True\n2.0");

class AppConfig{public static string Version{get;set;}="1.0";}

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
