// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_property_is_visible_through_interface
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

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

interface IValueHolder { int Value { get; } }
class Counter : IValueHolder {
    int IValueHolder.Value { get { return 12; } }
}
IValueHolder holder = new Counter();
__P((holder.Value).ToString());
__Check("12");
