// vybe-test: csharp/csharp_structs_value_semantics/struct_can_contain_reference_type_field
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

var wrapper = new Wrapper { Name = "text" }
;
__P((wrapper.Name).ToString());
__Check("text");

struct Wrapper { public string Name; }

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
