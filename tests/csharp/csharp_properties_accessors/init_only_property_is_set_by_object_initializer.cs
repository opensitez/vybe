// vybe-test: csharp/csharp_properties_accessors/init_only_property_is_set_by_object_initializer
// origin: languages/csharp/tests/csharp/test_csharp_properties_accessors.rs

using static __Harness;

var customer = new Customer { Name = "Ada", Tier = 3 }
;
__P((customer.Name).ToString());
__P((customer.Tier).ToString());
__Check("Ada\n3");

class Customer {
    public string Name { get; init; }
    public int Tier { get; init; }
}

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
