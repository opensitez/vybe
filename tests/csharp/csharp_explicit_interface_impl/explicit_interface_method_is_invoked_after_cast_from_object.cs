// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_is_invoked_after_cast_from_object
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

object item = new TaskRunner();
__P((((IRunner)item).Run()).ToString());
__Check("done");

interface IRunner { string Run(); }

class TaskRunner : IRunner {
    string IRunner.Run() { return "done"; }
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
