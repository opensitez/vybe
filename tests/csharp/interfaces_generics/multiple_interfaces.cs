// vybe-test: csharp/interfaces_generics/multiple_interfaces
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

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
var d = new Doc { Name = "test" };
d.Print();
__P((d.Serialize()).ToString());
__Check("Printing: test\nDOC:test");
