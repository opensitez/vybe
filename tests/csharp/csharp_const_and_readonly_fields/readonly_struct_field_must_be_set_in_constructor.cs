// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_struct_field_must_be_set_in_constructor
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

using static __Harness;

__P((new Cell(8).Value).ToString());
__Check("8");

struct Cell {
    public readonly int Value;
    public Cell(int value) { Value = value; }
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
