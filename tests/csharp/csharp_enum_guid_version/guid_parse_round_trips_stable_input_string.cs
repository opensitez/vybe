// vybe-test: csharp/csharp_enum_guid_version/guid_parse_round_trips_stable_input_string
// origin: languages/csharp/tests/csharp/test_csharp_enum_guid_version.rs

using static __Harness;

var text = "00112233-4455-6677-8899-aabbccddeeff";
__P((System.Guid.Parse(text).ToString()).ToString());
__Check("00112233-4455-6677-8899-aabbccddeeff");

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
