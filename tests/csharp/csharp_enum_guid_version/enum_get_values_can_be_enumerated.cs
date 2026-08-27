// vybe-test: csharp/csharp_enum_guid_version/enum_get_values_can_be_enumerated
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

foreach (var value in System.Enum.GetValues(typeof(State))) __P((value).ToString());
__Check("Idle\nRunning");

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
