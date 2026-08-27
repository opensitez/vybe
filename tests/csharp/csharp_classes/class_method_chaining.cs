// vybe-test: csharp/csharp_classes/class_method_chaining
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

using static __Harness;

var r = new Builder().Add("a").Add("b").Add("c").Build();
__P((r).ToString());
__Check("abc");

class Builder {
    private string result = "";
    public Builder Add(string s) {
        result += s;
        return this;
    }
    public string Build() { return result; }
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
