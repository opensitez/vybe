// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_instance_method_group_to_func
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

using static __Harness;

System.Func<int, int> fn = new Scale().Apply;
__P((fn(5)).ToString());
__Check("10");

class Scale { public int factor = 2; public int Apply(int n) => n * factor; }

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
