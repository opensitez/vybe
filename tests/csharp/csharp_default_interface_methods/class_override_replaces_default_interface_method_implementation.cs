// vybe-test: csharp/csharp_default_interface_methods/class_override_replaces_default_interface_method_implementation
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

using static __Harness;

IFormat fmt = new Custom();
__P((fmt.Label(3)).ToString());
__Check("x:3");

interface IFormat {
    string Format(int n);
    string Label(int n) { return "d:" + Format(n); }
}

class Custom : IFormat {
    public string Format(int n) { return n.ToString(); }
    public string Label(int n) { return "x:" + Format(n); }
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
