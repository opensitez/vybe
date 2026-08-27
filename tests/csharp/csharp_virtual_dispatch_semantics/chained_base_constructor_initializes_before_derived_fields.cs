// vybe-test: csharp/csharp_virtual_dispatch_semantics/chained_base_constructor_initializes_before_derived_fields
// origin: languages/csharp/tests/csharp/test_csharp_virtual_dispatch_semantics.rs

using static __Harness;

__P((new Child("id", "name").Read()).ToString());
__Check("id:name");

class Base {
    protected string token;
    public Base(string token) { this.token = token; }
}

class Child : Base {
    public string Label;
    public Child(string token, string label) : base(token) { Label = label; }
    public string Read() { return token + ":" + Label; }
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
