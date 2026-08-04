// vybe-test: csharp/csharp_interface_explicit_impl/class_with_both_explicit_and_public_overloads_picks_by_static_type
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

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

interface IDescribe { string Describe(); }
class Widget : IDescribe {
    public string Describe() => "widget";
    string IDescribe.Describe() => "interface:widget";
}
var w = new Widget();
IDescribe i = w;
__P((w.Describe()).ToString());
__P((i.Describe()).ToString());
__Check("widget\ninterface:widget");
