// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_interface_is_implemented_by_qualified_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

using static __Harness;

Demo.IRun worker = new Demo.Worker();
__P((worker.Run()).ToString());
__Check("done");

namespace Demo { public interface IRun { string Run(); } public class Worker : IRun { public string Run() { return "done"; } } }

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
