// vybe-test: csharp/csharp_file_scoped_namespace/file_scoped_namespace_field_mutation
// origin: languages/csharp/tests/csharp/test_csharp_file_scoped_namespace.rs

using static __Harness;

var c = new Counter { Count = 1 }
;
c.Count = 5;
__P((c.Count).ToString());
__Check("5");

class Counter { public int Count; }

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
