// vybe-test: csharp/interfaces_generics/multiple_interfaces
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

using static __Harness;

var d = new Doc { Name = "test" }
;
d.Print();
__P((d.Serialize()).ToString());
__Check("Printing: test\nDOC:test");

interface IPrintable {
    void Print();
}

interface ISerializable {
    string Serialize();
}

class Doc : IPrintable, ISerializable {
    public string Name;
    public void Print() { __P(("Printing: " + Name).ToString()); }
    public string Serialize() { return "DOC:" + Name; }
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
