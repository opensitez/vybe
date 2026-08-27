// vybe-test: csharp/csharp_with_expression_records/with_nested_replace_inner
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var p=new Person("Ann",new Address("Oslo"));
var q=p with{Home=new Address("Paris")}
;
__P((q.Home.City).ToString());
__Check("Paris");

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
