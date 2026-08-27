// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_percent_full
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

using static __Harness;

__P((new Tank().Percent).ToString());
__Check("75");

class Tank { public int level = 75; public int capacity = 100; public int Percent => level * 100 / capacity; }

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
