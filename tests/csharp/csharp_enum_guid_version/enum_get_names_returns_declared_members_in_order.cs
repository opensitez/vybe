// vybe-test: csharp/csharp_enum_guid_version/enum_get_names_returns_declared_members_in_order
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

foreach (var name in System.Enum.GetNames(typeof(State))) __P((name).ToString());
__Check("Idle\nRunning\nDone");

enum State { Idle, Running, Done }

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
