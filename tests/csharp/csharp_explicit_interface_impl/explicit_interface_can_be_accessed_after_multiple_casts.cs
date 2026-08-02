// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_can_be_accessed_after_multiple_casts
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICode { string Value(); }
class Ticket : ICode {
    string ICode.Value() { return "T-9"; }
}
var ticket = new Ticket();
object boxed = ticket;
__Check((((ICode)boxed).Value()).ToString(), "T-9");
