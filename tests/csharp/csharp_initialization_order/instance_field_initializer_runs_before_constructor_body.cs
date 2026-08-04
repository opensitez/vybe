// vybe-test: csharp/csharp_initialization_order/instance_field_initializer_runs_before_constructor_body
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

class Widget {
    string label = Init("field");
    public Widget() {
        __P(("ctor").ToString());
    }
    static string Init(string part) {
        __P((part).ToString());
        return part;
    }
}
new Widget();
__Check("field\nctor");
