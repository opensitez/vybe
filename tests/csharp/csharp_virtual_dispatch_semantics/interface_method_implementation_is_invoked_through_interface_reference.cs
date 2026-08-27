// vybe-test: csharp/csharp_virtual_dispatch_semantics/interface_method_implementation_is_invoked_through_interface_reference
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

IFormatter formatter = new DecimalFormatter();
__P((formatter.Format(4)).ToString());
__Check("004");

interface IFormatter {
    string Format(int value);
}

class DecimalFormatter : IFormatter {
    public string Format(int value) { return value.ToString("D3"); }
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
