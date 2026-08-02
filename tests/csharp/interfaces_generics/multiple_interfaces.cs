// vybe-test: csharp/interfaces_generics/multiple_interfaces
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
    public void Print() { __Check(("Printing: " + Name).ToString(), "Printing: test"); }
    public string Serialize() { return "DOC:" + Name; }
}
var d = new Doc { Name = "test" };
d.Print();
__Check((d.Serialize()).ToString(), "DOC:test");
