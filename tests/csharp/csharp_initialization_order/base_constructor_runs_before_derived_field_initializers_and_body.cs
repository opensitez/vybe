// vybe-test: csharp/csharp_initialization_order/base_constructor_runs_before_derived_field_initializers_and_body
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public Base() { __Check(("base-ctor").ToString(), "base-ctor"); }
}
class Derived : Base {
    string tag = Init("derived-field");
    public Derived() { __Check(("derived-ctor").ToString(), "derived-field"); }
    static string Init(string part) {
        __Check((part).ToString(), "derived-ctor");
        return part;
    }
}
new Derived();
