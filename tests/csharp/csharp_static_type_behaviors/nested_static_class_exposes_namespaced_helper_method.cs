// vybe-test: csharp/csharp_static_type_behaviors/nested_static_class_exposes_namespaced_helper_method
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

using static __Harness;

__P((TextTools.Parts.Join("a", "b")).ToString());
__Check("a/b");

class TextTools {
    public static class Parts {
        public static string Join(string a, string b) { return a + "/" + b; }
    }
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
