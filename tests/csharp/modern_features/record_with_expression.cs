// vybe-test: csharp/modern_features/record_with_expression
// origin: languages/csharp/tests/csharp/test_modern_features.rs

using static __Harness;

var p1 = new Person("Alice", 30);
var p2 = p1 with { Age = 31 }
;
__P((p1).ToString());
__P((p2).ToString());
__Check("Person { Name = Alice, Age = 30 }\nPerson { Name = Alice, Age = 31 }");

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
