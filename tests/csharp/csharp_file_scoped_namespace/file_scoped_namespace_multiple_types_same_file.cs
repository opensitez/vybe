// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_multiple_types_same_file
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

__P((new A().Value).ToString());
__P((new B().Value).ToString());
__Check("1\n2");

class A { public int Value = 1; }

class B { public int Value = 2; }

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
