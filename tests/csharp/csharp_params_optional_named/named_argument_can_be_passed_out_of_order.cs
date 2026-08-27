// vybe-test: csharp/csharp_params_optional_named/named_argument_can_be_passed_out_of_order
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

using static __Harness;

string Concat(string a, string b, string c) => a+b+c;
__P((Concat(c:"3",a:"1",b:"2")).ToString());
__Check("123");

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
