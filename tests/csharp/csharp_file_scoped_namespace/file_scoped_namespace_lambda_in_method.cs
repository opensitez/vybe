// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_lambda_in_method
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

__P((new Fn().Run()).ToString());
__Check("4");

class Fn { public int Run() { System.Func<int, int> f = x => x + 1; return f(3); } }

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
