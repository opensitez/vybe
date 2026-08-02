// vybe-test: csharp/csharp_explicit_interface_impl/explicit_implementations_disambiguate_same_method_name
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((((IText)payload).Format()).ToString(), "text");
__Check((((IJson)payload).Format()).ToString(), "json");
