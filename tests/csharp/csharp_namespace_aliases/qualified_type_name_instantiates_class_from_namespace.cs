// vybe-test: csharp/csharp_namespace_aliases/qualified_type_name_instantiates_class_from_namespace
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

var box = new Demo.Box();
__P((box.Name).ToString());
__Check("demo");

namespace Demo { public class Box { public string Name => "demo"; } }

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
