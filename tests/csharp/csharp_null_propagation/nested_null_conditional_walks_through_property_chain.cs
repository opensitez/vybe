// vybe-test: csharp/csharp_null_propagation/nested_null_conditional_walks_through_property_chain
// origin: languages/csharp/tests/csharp/test_csharp_null_propagation.rs

using static __Harness;

var user = new User { Address = new Address { City = "Paris" } }
;
__P((user?.Address?.City ?? "none").ToString());
__Check("Paris");

class Address { public string City { get; set; } }

class User { public Address Address { get; set; } }

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
