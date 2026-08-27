// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_operator_overload
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

var r = new V { N = 2 }
+ new V { N = 3 }
;
__P((r.N).ToString());
__Check("5");

struct V { public int N; public static V operator +(V a, V b) => new V { N = a.N + b.N }; }

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
