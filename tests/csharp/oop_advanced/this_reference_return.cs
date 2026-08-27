// vybe-test: csharp/oop_advanced/this_reference_return
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var b = new Builder();
__P((b.Add("A").Add("B").Add("C").Build()).ToString());
__Check("[A, B, C]");

class Builder {
    string parts = "";
    public Builder Add(string part) {
        if (parts.Length > 0) parts += ", ";
        parts += part;
        return this;
    }
    public string Build() { return "[" + parts + "]"; }
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
