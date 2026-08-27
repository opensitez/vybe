// vybe-test: csharp/csharp_string_interpolation/simple_variable_interpolation_in_string
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

using static __Harness;

string name="World";
__P(($"Hello {name}!").ToString());
__Check("Hello World!");

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
