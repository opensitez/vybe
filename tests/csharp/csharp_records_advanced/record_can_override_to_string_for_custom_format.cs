// vybe-test: csharp/csharp_records_advanced/record_can_override_to_string_for_custom_format
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

__P((new User("Ada")).ToString());
__Check("User:Ada");

record User(string Name) { public override string ToString() { return $"User:{Name}"; } }

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
