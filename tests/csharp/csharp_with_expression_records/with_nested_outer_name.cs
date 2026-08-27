// vybe-test: csharp/csharp_with_expression_records/with_nested_outer_name
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var q=(new Person("Ann",new Address("Oslo"))) with{Name="Bob"}
;
__P((q.Name).ToString());
__Check("Bob");

record Address(string City);

record Person(string Name,Address Home);

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
