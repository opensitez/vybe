// vybe-test: csharp/csharp_default_interface_methods_deep/diamond_three_interfaces_class_unified_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

using static __Harness;

IDemoService svc = new DemoServiceImpl();
__P(svc.GetGreeting());
__Check("Hello_diamond_three_interfaces_class_unified_override");

interface IDemoService {
    string GetGreeting() => "Hello_diamond_three_interfaces_class_unified_override";
}
class DemoServiceImpl : IDemoService { }
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
