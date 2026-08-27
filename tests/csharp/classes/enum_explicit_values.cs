// vybe-test: csharp/classes/enum_explicit_values
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

__P((Status.Ok).ToString());
__P((Status.NotFound).ToString());
__Check("Ok\nNotFound");

enum Status { Ok = 200, NotFound = 404, Error = 500 }

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
