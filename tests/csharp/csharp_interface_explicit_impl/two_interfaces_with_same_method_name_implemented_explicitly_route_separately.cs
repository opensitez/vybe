// vybe-test: csharp/csharp_interface_explicit_impl/two_interfaces_with_same_method_name_implemented_explicitly_route_separately
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

interface ILeft  { string Side(); }
interface IRight { string Side(); }
class Both : ILeft, IRight {
    string ILeft.Side()  => "left";
    string IRight.Side() => "right";
}
ILeft  l = new Both();
IRight r = new Both();
__P((l.Side()).ToString());
__P((r.Side()).ToString());
__Check("left\nright");
