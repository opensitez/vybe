// vybe-test: csharp/csharp_exception_types/divide_by_zero_exception_thrown_by_integer_division
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

using static __Harness;

string result = "";
try { int x = 10 / Math.Max(0, 0); }
catch(System.DivideByZeroException e) { result = "dbz"; }
__P((result).ToString());
__Check("dbz");

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
