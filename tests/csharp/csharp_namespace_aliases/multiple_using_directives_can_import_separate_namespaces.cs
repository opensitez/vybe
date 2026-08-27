// vybe-test: csharp/csharp_namespace_aliases/multiple_using_directives_can_import_separate_namespaces
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;
using Demo.Left;
using Demo.Right;

__P((new A().Name + new B().Name).ToString());
__Check("AB");

namespace Demo.Left { public class A { public string Name => "A"; } }

namespace Demo.Right { public class B { public string Name => "B"; } }

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
