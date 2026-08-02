// vybe-test: csharp/csharp_properties_accessors/read_only_property_is_assigned_from_constructor
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class BuildInfo {
    public string Version { get; }
    public BuildInfo(string version) { Version = version; }
}
var info = new BuildInfo("1.2.3");
__Check((info.Version).ToString(), "1.2.3");
