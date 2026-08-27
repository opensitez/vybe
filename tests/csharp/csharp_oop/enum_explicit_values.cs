// vybe-test: csharp/csharp_oop/enum_explicit_values
// origin: languages/csharp/tests/csharp/test_csharp_oop.rs

using static __Harness;

__P(((int)HttpStatus.NotFound).ToString());
__Check("404");

enum HttpStatus {
    OK = 200,
    NotFound = 404,
    ServerError = 500
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
