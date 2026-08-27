// vybe-test: csharp/csharp_enum_guid_version/enum_try_parse_reports_success_for_valid_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

System.Enum.TryParse<State>("Idle", out var value);
__P((value).ToString());
__Check("Idle");

enum State { Idle, Running }

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
