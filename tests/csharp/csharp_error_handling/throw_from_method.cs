// vybe-test: csharp/csharp_error_handling/throw_from_method
// origin: languages/csharp/tests/csharp/test_csharp_error_handling.rs

using static __Harness;

int Divide(int a, int b) {
    if (b == 0) throw new Exception("Division by zero");
    return a / b;
}
try {
    __P((Divide(10, 2)).ToString());
    __P((Divide(10, 0)).ToString());
}
catch (Exception e) {
    __P((e.Message).ToString());
}
__Check("5\nDivision by zero");

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
