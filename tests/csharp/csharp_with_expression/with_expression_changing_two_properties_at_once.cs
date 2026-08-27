// vybe-test: csharp/csharp_with_expression/with_expression_changing_two_properties_at_once
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

using static __Harness;

var p = new Person("Ada", 30);
var updated = p with { Name = "Grace", Age = 31 }
;
__P((updated.Name).ToString());
__P((updated.Age).ToString());
__Check("Grace\n31");

record Person(string Name, int Age);

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
