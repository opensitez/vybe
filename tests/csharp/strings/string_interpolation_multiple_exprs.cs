// vybe-test: csharp/strings/string_interpolation_multiple_exprs
// origin: languages/csharp/tests/csharp/test_strings.rs

using static __Harness;

var a = "Alice";
var age = 30;
__P(($"{a} is {age} years old").ToString());
__Check("Alice is 30 years old");

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
