// vybe-test: csharp/csharp_properties_advanced/lazy_initialized_property_created_on_first_access
// origin: languages/csharp/tests/csharp/test_csharp_properties_advanced.rs

using static __Harness;

var c=new Config();
__P((c.Tag).ToString());
__Check("computed");

class Config{
    System.Lazy<string> _tag=new System.Lazy<string>(()=>"computed");
    public string Tag=>_tag.Value;
}

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
