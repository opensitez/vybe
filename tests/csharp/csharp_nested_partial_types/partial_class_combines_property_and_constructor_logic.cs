// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_property_and_constructor_logic
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

__P((new Build("nightly").Name).ToString());
__Check("nightly");

partial class Build {
    public string Name { get; set; }
}

partial class Build {
    public Build(string name) { Name = name; }
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
