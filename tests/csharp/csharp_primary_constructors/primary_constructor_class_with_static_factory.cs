// vybe-test: csharp/csharp_primary_constructors/primary_constructor_class_with_static_factory
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

using static __Harness;

__P((Token.Default().Id).ToString());
__Check("0");

class Token(int id) {
    public int Id => id;
    public static Token Default() => new Token(0);
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
