// vybe-test: csharp/csharp_structs_value_semantics/nullable_struct_value_exposes_underlying_member
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

using static __Harness;

System.DateTime? value = new System.DateTime(2024, 1, 1);
__P((value.Value.Year).ToString());
__Check("2024");

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
