// vybe-test: csharp/csharp_checked_overflow_throws/checked_add_on_int_max_value_throws_overflow_exception
// origin: languages/csharp/tests/csharp/test_csharp_checked_overflow_throws.rs

using static __Harness;

string outcome = "ok";
try {
    checked {
        int value = int.MaxValue;
        value += 1;
    }
}
catch (System.OverflowException) {
    outcome = "overflow";
}
__P((outcome).ToString());
__Check("overflow");

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
