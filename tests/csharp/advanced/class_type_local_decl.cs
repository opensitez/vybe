// vybe-test: csharp/advanced/class_type_local_decl
// origin: languages/csharp/tests/csharp/test_advanced.rs

using static __Harness;

Foo f = new Foo("hello");
__P((f.name).ToString());
__Check("hello");

class Foo {
            public string name;
            public Foo(string n) { this.name = n; }
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
