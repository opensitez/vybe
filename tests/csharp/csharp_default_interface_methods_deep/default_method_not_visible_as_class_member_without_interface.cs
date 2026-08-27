// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_not_visible_as_class_member_without_interface
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

using static __Harness;

IDemoService svc = new DemoServiceImpl();
__P(svc.GetGreeting());
__Check("Hello_default_method_not_visible_as_class_member_without_interface");

interface IDemoService {
    string GetGreeting() => "Hello_default_method_not_visible_as_class_member_without_interface";
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
