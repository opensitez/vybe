// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_can_be_accessed_after_multiple_casts
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

interface ICode { string Value(); }
class Ticket : ICode {
    string ICode.Value() { return "T-9"; }
}
var ticket = new Ticket();
object boxed = ticket;
__P((((ICode)boxed).Value()).ToString());
__Check("T-9");
