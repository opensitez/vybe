// vybe-test: csharp/csharp_enum_guid_version/enum_parse_converts_string_to_enum_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

__P((System.Enum.Parse(typeof(State), "Running")).ToString());
__Check("Running");

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
