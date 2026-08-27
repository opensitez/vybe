// vybe-test: csharp/csharp_collection_initializer_syntax/nested_object_initializer_builds_graph_in_one_expression
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

using static __Harness;

var person = new Person { Home = new Address { City = "Oslo" } }
;
__P((person.Home.City).ToString());
__Check("Oslo");

class Address { public string City { get; set; } }

class Person { public Address Home { get; set; } }

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
