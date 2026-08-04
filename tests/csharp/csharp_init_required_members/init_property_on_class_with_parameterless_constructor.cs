// vybe-test: csharp/csharp_init_required_members/init_property_on_class_with_parameterless_constructor
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

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

class Widget { public Widget() { } public int Count { get; init; } = 0; }
var w = new Widget { Count = 5 };
__P((w.Count).ToString());
__Check("5");
