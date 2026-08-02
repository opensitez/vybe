// vybe-test: csharp/csharp_initialization_order/instance_field_initializer_runs_before_constructor_body
// origin: languages/csharp/tests/csharp/test_csharp_initialization_order.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget {
    string label = Init("field");
    public Widget() {
        __Check(("ctor").ToString(), "field");
    }
    static string Init(string part) {
        __Check((part).ToString(), "ctor");
        return part;
    }
}
new Widget();
