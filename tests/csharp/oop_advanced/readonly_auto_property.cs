// vybe-test: csharp/oop_advanced/readonly_auto_property
// origin: languages/csharp/tests/csharp/test_oop_advanced.rs

using static __Harness;

var p = new Person("Alice");
__P((p.Name).ToString());
__Check("Alice");

class Person {
    public string Name { get; }
    public Person(string name) { Name = name; }
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
