// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_struct_field_access
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

var p = new Point { X = 2, Y = 3 }
;
__P((p.X + p.Y).ToString());
__Check("5");

struct Point { public int X; public int Y; }

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
