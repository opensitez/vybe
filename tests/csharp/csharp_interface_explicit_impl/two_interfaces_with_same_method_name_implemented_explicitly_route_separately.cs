// vybe-test: csharp/csharp_interface_explicit_impl/two_interfaces_with_same_method_name_implemented_explicitly_route_separately
// origin: languages/csharp/tests/csharp/test_csharp_interface_explicit_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((l.Side()).ToString(), "left");
__Check((r.Side()).ToString(), "right");
