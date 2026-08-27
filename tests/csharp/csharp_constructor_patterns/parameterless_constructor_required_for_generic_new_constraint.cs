// vybe-test: csharp/csharp_constructor_patterns/parameterless_constructor_required_for_generic_new_constraint
// origin: languages/csharp/tests/csharp/test_csharp_constructor_patterns.rs

using static __Harness;

T Make<T>() where T:new()=>new T();
__P((Make<Widget>().Value).ToString());
__Check("7");

class Widget{public int Value=7;}

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
