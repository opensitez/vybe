// vybe-test: csharp/csharp_generics_advanced/generic_method_infers_type_from_argument
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

T Identity<T>(T value) => value;
__P((Identity(99)).ToString());
__P((Identity("hi")).ToString());
__Check("99\nhi");

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
