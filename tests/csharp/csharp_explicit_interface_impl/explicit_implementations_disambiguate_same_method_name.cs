// vybe-test: csharp/csharp_explicit_interface_impl/explicit_implementations_disambiguate_same_method_name
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

var payload = new Payload();
__P((((IText)payload).Format()).ToString());
__P((((IJson)payload).Format()).ToString());
__Check("text\njson");

interface IText { string Format(); }

interface IJson { string Format(); }

class Payload : IText, IJson {
    string IText.Format() { return "text"; }
    string IJson.Format() { return "json"; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
