// vybe-test: csharp/csharp_reflection/get_methods_count_includes_public_instance_methods
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

using static __Harness;

__P((typeof(Calc).GetMethods(
    System.Reflection.BindingFlags.Public|System.Reflection.BindingFlags.Instance|
    System.Reflection.BindingFlags.DeclaredOnly).Length).ToString());
__Check("2");

class Calc { public int Add(int a, int b) => a+b; public int Sub(int a, int b) => a-b; }

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
