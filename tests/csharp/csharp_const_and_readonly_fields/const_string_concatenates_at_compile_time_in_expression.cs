// vybe-test: csharp/csharp_const_and_readonly_fields/const_string_concatenates_at_compile_time_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

using static __Harness;

__P((Labels.Full).ToString());
__Check("user_id");

class Labels {
    public const string Base = "user";
    public const string Full = Base + "_id";
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
