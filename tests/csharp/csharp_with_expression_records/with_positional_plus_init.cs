// vybe-test: csharp/csharp_with_expression_records/with_positional_plus_init
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var v=(new User("Ada"){Age=20}) with{Age=21}
;
__P((v.Age).ToString());
__Check("21");

record User(string Name){public int Age{get;init;}}

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
