// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_bool_is_empty
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Bag { Data = "" }.IsEmpty).ToString());
__P((new Bag { Data = "x" }.IsEmpty).ToString());
__Check("True\nFalse");

class Bag { public string? Data; public bool IsEmpty => Data == null || Data.Length == 0; }

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
