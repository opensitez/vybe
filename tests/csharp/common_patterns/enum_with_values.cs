// vybe-test: csharp/common_patterns/enum_with_values
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

using static __Harness;

__P(((int)Status.Active).ToString());
__P(((int)Status.Inactive).ToString());
__P(((int)Status.Pending).ToString());
__Check("1\n0\n2");

enum Status { Active = 1, Inactive = 0, Pending = 2 }

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
