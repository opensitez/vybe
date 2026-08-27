// vybe-test: csharp/csharp_init_required_members/init_property_byte_value_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

using static __Harness;

int val = 100;
__P(val.ToString());
__Check("100");
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
