// vybe-test: csharp/csharp_record_types/records_generated_tostring_includes_property_names_and_values
// origin: languages/csharp/tests/csharp/test_csharp_record_types.rs

using static __Harness;

__P((new Tag("admin").ToString().Contains("admin")).ToString());
__Check("True");

record Tag(string Name);

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
