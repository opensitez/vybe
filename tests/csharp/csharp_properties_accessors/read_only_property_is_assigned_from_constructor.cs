// vybe-test: csharp/csharp_properties_accessors/read_only_property_is_assigned_from_constructor
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var info = new BuildInfo("1.2.3");
__P((info.Version).ToString());
__Check("1.2.3");

class BuildInfo {
    public string Version { get; }
    public BuildInfo(string version) { Version = version; }
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
