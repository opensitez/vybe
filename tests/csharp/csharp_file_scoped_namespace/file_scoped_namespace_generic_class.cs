// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_generic_class
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

var b = new Box<int> { Item = 9 }
;
__P((b.Item).ToString());
__Check("9");

class Box<T> { public T Item; }

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
