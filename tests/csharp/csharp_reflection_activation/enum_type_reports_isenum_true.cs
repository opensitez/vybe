// vybe-test: csharp/csharp_reflection_activation/enum_type_reports_isenum_true
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

using static __Harness;

__P((typeof(State).IsEnum).ToString());
__Check("True");

enum State { Ready }

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
