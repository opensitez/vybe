// vybe-test: csharp/csharp_interface_explicit_impl/class_with_both_explicit_and_public_overloads_picks_by_static_type
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((w.Describe()).ToString(), "widget");
__Check((i.Describe()).ToString(), "interface:widget");
