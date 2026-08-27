// vybe-test: csharp/csharp_records_advanced/record_with_non_positional_members_can_use_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

using static __Harness;

var theme = new Theme { Name = "light", Version = 2 }
;
__P((theme.Name + ":" + theme.Version).ToString());
__Check("light:2");

record Theme { public string Name { get; init; } public int Version { get; init; } }

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
