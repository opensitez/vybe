// vybe-test: csharp/csharp_nested_type_member_access/nested_static_class_reads_outer_static_private_state
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_member_access.rs

using static __Harness;

__P((Outer.Via()).ToString());
__Check("3");

class Outer {
    static int tally = 3;
    static class Inner {
        public static int Read() { return tally; }
    }
    public static int Via() { return Inner.Read(); }
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
