// vybe-test: csharp/csharp_nested_classes/nested_static_class_provides_utility_methods
// origin: languages/csharp/tests/csharp/test_csharp_nested_classes.rs

using static __Harness;

__P((Parser.Helpers.ToInt("99")).ToString());
__Check("99");

class Parser{
    public static class Helpers{public static int ToInt(string s)=>int.Parse(s);}
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
