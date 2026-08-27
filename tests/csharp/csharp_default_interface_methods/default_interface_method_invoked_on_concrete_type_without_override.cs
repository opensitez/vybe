// vybe-test: csharp/csharp_default_interface_methods/default_interface_method_invoked_on_concrete_type_without_override
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

using static __Harness;

IBanner banner = new ConsoleReporter();
__P(banner.Banner());
__Check("ReportBanner");

interface IBanner {
    string Banner() => "ReportBanner";
}
class ConsoleReporter : IBanner { }
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
