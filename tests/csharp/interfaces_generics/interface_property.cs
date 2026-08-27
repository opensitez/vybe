// vybe-test: csharp/interfaces_generics/interface_property
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

INamed p = new Person { Name = "Alice" }
;
__P((p.Name).ToString());
__Check("Alice");

interface INamed {
    string Name { get; }
}

class Person : INamed {
    public string Name { get; set; }
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
