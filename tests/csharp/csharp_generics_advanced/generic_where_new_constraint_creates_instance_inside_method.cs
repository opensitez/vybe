// vybe-test: csharp/csharp_generics_advanced/generic_where_new_constraint_creates_instance_inside_method
// origin: languages/csharp/tests/csharp/test_csharp_generics_advanced.rs

using static __Harness;

T Make<T>() where T : new() => new T();
__P((Make<Widget>().Val).ToString());
__Check("5");

class Widget { public int Val = 5; }

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
