// vybe-test: csharp/csharp_with_expression_records/with_nested_triple_inner
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

using static __Harness;

var p=new Person("A",new Address("Oslo",new Zip("01")));
var q=p with{Home=p.Home with{Z=p.Home.Z with{Code="02"}}}
;
__P((q.Home.Z.Code).ToString());
__Check("02");

record Zip(string Code);

record Address(string City,Zip Z);

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
