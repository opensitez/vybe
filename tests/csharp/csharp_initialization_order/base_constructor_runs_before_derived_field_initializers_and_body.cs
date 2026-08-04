// vybe-test: csharp/csharp_initialization_order/base_constructor_runs_before_derived_field_initializers_and_body
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Base {
    public Base() { __P(("base-ctor").ToString()); }
}
class Derived : Base {
    string tag = Init("derived-field");
    public Derived() { __P(("derived-ctor").ToString()); }
    static string Init(string part) {
        __P((part).ToString());
        return part;
    }
}
new Derived();
__Check("base-ctor\nderived-field\nderived-ctor");
