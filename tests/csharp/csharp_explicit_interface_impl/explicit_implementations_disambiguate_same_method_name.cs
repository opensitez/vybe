// vybe-test: csharp/csharp_explicit_interface_impl/explicit_implementations_disambiguate_same_method_name
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

interface IText { string Format(); }
interface IJson { string Format(); }
class Payload : IText, IJson {
    string IText.Format() { return "text"; }
    string IJson.Format() { return "json"; }
}
var payload = new Payload();
__P((((IText)payload).Format()).ToString());
__P((((IJson)payload).Format()).ToString());
__Check("text\njson");
