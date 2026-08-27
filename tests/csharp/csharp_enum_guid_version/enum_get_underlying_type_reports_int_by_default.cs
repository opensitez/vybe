// vybe-test: csharp/csharp_enum_guid_version/enum_get_underlying_type_reports_int_by_default
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

__P((System.Enum.GetUnderlyingType(typeof(State)).Name).ToString());
__Check("Int32");

enum State { Idle }

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
