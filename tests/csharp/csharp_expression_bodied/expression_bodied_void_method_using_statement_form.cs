// vybe-test: csharp/csharp_expression_bodied/expression_bodied_void_method_using_statement_form
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied.rs

using static __Harness;

new Logger().Log("hello");
__Check("hello");

class Logger{public void Log(string msg)=>__P((msg).ToString());}

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
